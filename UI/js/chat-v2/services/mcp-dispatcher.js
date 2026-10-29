/**
 * McpDispatcher — 确定性 MCP 工具调度层
 * 
 * 问题背景：
 *   Ollama 等本地小模型（7B 级）在 chat-only 模式下根本不会输出工具标签。
 *   不管是 Excel/PPT/Computer 还是未来的 Browser/WPS，所有纯 MCP 工具
 *   丢给 Ollama 都是幻觉或空回复——重试多少次都一样。
 * 
 * 解决方案：
 *   在 Ollama 之前加一层确定性匹配：
 *   1) 后端 McpCapabilityRegistry.Match()（最准）
 *   2) 前端 toolRegistry ngram+keywords 兜底（bridge 不可用时）
 *   命中 → 直接 invokeTracked 执行，100ms 内完成
 *   未命中 → 原流程（Skill 路由 → Ollama → agent-loop）
 * 
 * 架构统一：
 *   以后加 BrowserAgent / WpsAgent 等新 MCP 类型，零额外维护。
 *   只要工具注册进 registry（带 description + keywords），
 *   调度层自动识别、自动匹配、自动执行。
 */

if (!window.App) window.App = {};
if (!window.App.ChatV2) window.App.ChatV2 = {};

const ChatLog = window.ChatLog || console;

// 路由常量访问器：统一从 routing-constants（单一事实源）取硬短路触发关键词，
// 避免在本文件内联重复定义导致"改一处漏一处"漂移。
const RC = () => ((window.App && App.ChatV2 && App.ChatV2.routingConstants) || {});

// ── 单元格格式化动词 → excel_format_range 参数（确定性硬短路辅助）──
// 从用户消息抽取显式区域（A1:D1 / $A$1:$D$1 / Sheet1!A1:D1 / 单格 A1）。
// 仅当命中明确区域时才使用，绝不静默用当前选区顶替（选区兜底在 FormatRange.vb 后端，
// 会把"没解析到区域"误变成"格式化整个选区"，已在多次测试中造成做错区域）。
function extractExcelRange(msg) {
    if (typeof msg !== 'string') return null;
    // 负向后顾/前瞻：避免把 Sheet1 里的 "t1"、或 "A1产品" 误当成区域。
    const re = /(?<![A-Za-z0-9_])(?:([^!:\s，。！？、]{1,31})!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?:\s*[:：]\s*\$?([A-Za-z]{1,3})\$?(\d{1,7}))?(?![A-Za-z0-9_])/;
    const m = msg.match(re);
    if (!m) return null;
    const sheet = m[1] || '';
    const c1 = m[2], r1 = m[3], c2 = m[4], r2 = m[5];
    const range = c2 ? (c1 + r1 + ':' + c2 + r2) : (c1 + r1);
    return (sheet ? sheet + '!' : '') + range;
}

// 把常见格式动词映射为 excel_format_range 的 format/alignment 对象。
// 仅当识别出至少一个具体样式信号时才返回非空，避免"格式化"这类无细节请求被硬调成预设。
function buildFormatArgs(message) {
    const m_ = (message || '').toLowerCase();
    const format = {};
    const alignment = {};
    let signal = false;
    if (m_.includes('加粗')) { format.bold = true; signal = true; }
    if (m_.includes('倾斜') || m_.includes('斜体')) { format.italic = true; signal = true; }
    if (m_.includes('水平居中') || m_.includes('居中')) alignment.horizontal = 'center';
    if (m_.includes('垂直居中') || m_.includes('上下居中')) alignment.vertical = 'middle';
    if (m_.includes('边框')) { format.borderStyle = 'thin'; format.borderColor = '000000'; signal = true; }
    if (m_.includes('自动换行') || m_.includes('换行')) { format.wrapText = true; signal = true; }
    const fs = m_.match(/字号?\s*(\d{1,3})|(\d{1,3})\s*号字/);
    if (fs) { format.fontSize = parseInt(fs[1] || fs[2], 10); signal = true; }
    // 字体名用已知白名单匹配，避免把"字体颜色红"误当成字体名（与下方 fontColor 冲突）
    const KNOWN_FONTS = ['微软雅黑', '等线', '宋体', '黑体', '楷体', '仿宋', '新宋体', '幼圆', '华文行楷', 'times new roman', 'arial', 'calibri', 'cambria', 'georgia'];
    for (const f of KNOWN_FONTS) { if (m_.includes(f)) { format.fontName = f; signal = true; break; } }
    const colorOf = (s) => {
        const hex = s.match(/#?([0-9a-fA-F]{6})/);
        if (hex) return hex[1].toUpperCase();
        const names = { '红': 'FF0000', '蓝': '0000FF', '绿': '00FF00', '黄': 'FFFF00', '黑': '000000', '白': 'FFFFFF', '灰': '808080', '橙': 'FFA500', '紫': '800080', '粉': 'FFC0CB' };
        for (const k in names) if (s.indexOf(k) >= 0) return names[k];
        return null;
    };
    const fc = m_.match(/字体颜色[：: ]?\s*([^\s，。！？、,]+)/);
    if (fc) { const c = colorOf(fc[1]); if (c) { format.fontColor = c; signal = true; } }
    const bc = m_.match(/(?:填充色|背景色|底色)[：: ]?\s*([^\s，。！？、,]+)/);
    if (bc) { const c = colorOf(bc[1]); if (c) { format.bgColor = c; signal = true; } }
    if (!signal) return null;
    const args = {};
    if (Object.keys(format).length) args.format = format;
    if (Object.keys(alignment).length) args.alignment = alignment;
    return args;
}

/** 「美化表格」没有指定颜色、边框、字号。交给后端商务蓝预设，避免整表涂成一个浅色后声称已美化。 */
function isBareTableBeautify(message) {
    const s = String(message || '').trim().replace(/[。！？!?\s]+$/g, '');
    if (!s || s.length > 18) return false;
    if (/红|橙|黄|绿|青|蓝|紫|粉|灰|黑|白|棕|金|边框|字号|字体|加粗|居中|数据条|色阶|底色|背景|填充|列宽/.test(s)) return false;
    return /^(?:请|帮我|帮忙)?(?:把|将)?(?:当前|这个|这份|选中的|选区)?(?:表格|工作表|表|数据)?美化(?:一下|下)?(?:当前|这个|这份)?(?:表格|工作表|表)?$/.test(s)
        || /^(?:请|帮我|帮忙)?美化(?:一下|下)?(?:当前|这个|这份|选中的)?(?:表格|工作表|表|数据)?$/.test(s);
}

/**
 * 【多任务规划·格式化意图识别】识别消息里的美化/格式化诉求，并判断能否被确定性参数全覆盖。
 * 两处消费：
 *   ① 拆分硬短路："按X拆分成不同Sheet，然后分别对工作表美化"——第二项任务进同一条链
 *      （实测 2026-10-22：旧实现只跑拆分就宣称任务完成，9 个新表一个都没美化）；
 *   ② 批量美化硬短路："美化工作簿中的所有工作表"——枚举可见表后逐表美化
 *      （实测 2026-10-22：该请求落给 LLM 后只调 1 次 excel_format_range 且 range 填成 A1
 *       只美化 1 个单元格，随后用"正在依次应用相同预设至其他工作表，请稍候"充数，回合以
 *       "任务完成"收场——N 张表只做了 1 张甚至 1 个单元格）。
 *
 * 返回值：
 *   null         —— 没有美化/格式化诉求
 *   { args }     —— 可确定性覆盖：args={} 表示走后端商务蓝预设；带 format/alignment 表示用户点名了具体样式
 *   'uncovered'  —— 用户点名了具体样式但确定性参数覆盖不全（行高/底纹/对齐/数据条…）→ 调用方应放弃
 *                   硬短路，放行 agent-loop 多步处理（与格式化硬短路的"多要求排他"同一范式）
 */
function detectFormatIntent(message) {
    const m = String(message || '');
    if (!m) return null;
    // 否定/豁免："不要美化"/"保持格式不变"/"不要转为商务蓝" → 不当格式化意图。
    // ⚠ 不能用裸"别"： "分别对工作表美化"（分别= Positive!）会被 别.{0,8}美化 误判成否定。
    if (/(不要|无需|不用|不必|禁止|请勿|切勿|无须).{0,8}(美化|格式化|排版|样式|格式|主题|配色|风格|商务蓝|科技绿|活力橙|优雅紫|极简灰|金融红|中国红|朱红|国旗红|墨黑金|晴空青)|(格式|样式)不变|保持原样|维持原样|保持原有|维持原有|保持原来|保留原(?:格式|样式|样子)|格式不动|别动格式/.test(m)) return null;
    // 美化/格式化诉求：既认笼统的"美化/格式化/排版/样式"，也认具体样式动词与主题词
    //（"然后加粗"/"转为商务蓝"/"科技绿主题"），与格式化硬短路共用同一套判定词汇。
    // 数据条/色阶/图标集/条件格式也算"格式化诉求"（见下方豁免）——不列这里会漏判。
    const TRAILING_FMT = /(美化|格式化|排版|样式|加粗|加斜|倾斜|斜体|居中|对齐|边框|字号|字体|颜色|底色|背景色|填充|底纹|行高|列宽|换行|冻结|高亮|数据条|色阶|图标集|条件格式|主题|配色|风格|色系|商务蓝|科技绿|活力橙|优雅紫|极简灰|金融红|中国红|朱红|国旗红|墨黑金|晴空青|深蓝|黑金)/;
    if (!TRAILING_FMT.test(m)) return null;
    // 【清除类排除·2026-10-22】"清除所有工作表的样式/去掉其他表的背景色"是【删除格式】诉求，
    // 绝不能被 beautify 链接手反过来涂一层预设（实测：弱模型甚至把"美化"执行成 7 次
    // "清除现有格式"）。清除/还原类动词 + 格式名词同现 → 不是美化意图。
    if (/(清除|去掉|去除|删除|移除|重置|还原|恢复|清空).{0,10}(格式|样式|底色|背景|底纹|颜色|边框|填充)/.test(m)) return null;
    // 条件格式类（数据条/色阶/图标集）确定性链盖不住 → 直接放行，不拿预设冒充
    if (/(数据条|色阶|图标集|条件格式)/.test(m)) return 'uncovered';
    const fmt = buildFormatArgs(m);
    // 覆盖度判定同时管"解析出参数的"和"只有一个笼统词的"：笼统词里的行高/对齐/底纹
    // buildFormatArgs 解析不出，uncoveredFormatAspects 能点名 → 放行 agent-loop
    const miss = (typeof uncoveredFormatAspects === 'function') ? uncoveredFormatAspects(m, fmt || {}) : [];
    if (miss.length > 0) return 'uncovered';
    if (fmt) return { args: fmt, label: '指定样式' };      // 显式参数全覆盖 → 原样透传
    // 主题类："转为商务蓝"= 后端默认预设（深蓝表头白字，最完整）；其它主题 =
    // 表头浅色填充 + 主题色边框（excel_format_range 的 fontColor 作用于整个区域，
    // 没有"仅表头字色"参数，深底白字做不了，故用浅底+主题边框表达主题）。
    const theme = detectTheme(m);
    if (theme) {
        if (theme.preset) return { args: {}, label: '商务蓝美化预设' };
        return {
            args: {
                format: {
                    bgColor: theme.fill,
                    bgScope: 'header',
                    bold: true,
                    borderStyle: 'thin',
                    borderColor: theme.border
                }
            },
            label: theme.label + '主题（表头填充+主题色边框）'
        };
    }
    return { args: {}, label: '商务蓝美化预设' };   // 纯"美化" → 后端商务蓝专业预设
}

/**
 * 是否"批量美化"请求：一句话要对【多张工作表】做美化/格式化。
 * 判定 = 批量目标词（所有/全部/每个/各个/每张/批量/其他/其余/剩余/剩下/整个工作簿）+ 工作表名词
 *        + 美化/格式化诉求。
 * 格式诉求词含主题词（商务蓝/科技绿/主题/配色…）——实测 2026-10-22"将工作簿中的所有工作表
 * 转为商务蓝"因主题词不在表内被漏判，又变成"只说不做"。
 * 【2026-10-22 晚】"其他/其余/剩余/剩下"入表：实测"美化…其他工作表"（拆分/合并后对产物表
 * 美化的自然说法）不在旧表内 → 落给弱模型 → 被执行成 7 次"清除现有格式"（破坏性！第一个表
 * 字色与底色一致文字隐形）。有格式诉求词门控，"把结果写到其他工作表"（无格式词）不受影响。
 * 刻意不收"把每个表的数据按金额排序"这类非格式化复合请求（无格式词 → false，放行下游）。
 */
function isBatchBeautifyMessage(message) {
    const m = String(message || '');
    if (!m) return false;
    const BATCH_TARGET = /(所有|全部|各个|每个|每一张|每张|批量|其他|其余|剩余|剩下)/;
    const SHEET_NOUN = /(工作表|sheet|表格|数据表)/i;
    const WHOLE_BOOK = /(整个|整本|全部).{0,4}工作簿/i;
    if (!((BATCH_TARGET.test(m) && SHEET_NOUN.test(m)) || WHOLE_BOOK.test(m))) return false;
    return /(美化|格式化|排版|样式|加粗|居中|边框|字号|字体|颜色|底色|行高|列宽|对齐|换行|冻结|高亮|主题|配色|风格|色系|商务蓝|科技绿|活力橙|优雅紫|极简灰|金融红|墨黑金|晴空青|深蓝|黑金)/.test(m);
}

/**
 * Excel 主题色映射（值与 config/table-style-presets.js 的 PALETTES 对齐：
 * headerFill 浅色表头底 + border 主题色网格线）。
 * 为何用浅色表头底而不用 headerDeep：excel_format_range 的 fontColor 作用于**整个区域**，
 * 深底+黑字会不可读，而它没有"仅表头字色"的参数（bgScope 只管底色）——
 * 深蓝底白字那一套由后端商务蓝预设（空 format）负责，见 detectFormatIntent。
 * @returns {{fill:string, border:string, label:string, preset?:boolean}|null}
 */
function detectTheme(message) {
    const m = String(message || '');
    if (!m) return null;
    const THEMES = [
        { re: /商务蓝|深蓝/, fill: 'D9E2F3', border: '4472C4', label: '商务蓝' },
        { re: /科技绿/, fill: 'D9F2EA', border: '2A9D8F', label: '科技绿' },
        { re: /活力橙/, fill: 'FCE4D6', border: 'E76F51', label: '活力橙' },
        { re: /优雅紫/, fill: 'E5D9F2', border: '7030A0', label: '优雅紫' },
        { re: /极简灰/, fill: 'EDEDED', border: 'BFBFBF', label: '极简灰' },
        { re: /金融红/, fill: 'F8D7DA', border: 'C00000', label: '金融红' },
        { re: /墨黑金|黑金/, fill: 'EFE7D6', border: 'B99A4B', label: '墨黑金' },
        { re: /晴空青/, fill: 'D8EEF8', border: '2E9BC4', label: '晴空青' }
    ];
    for (const t of THEMES) {
        if (t.re.test(m)) {
            // 商务蓝本身就是后端默认预设（深蓝表头+白字+细边框+自动列宽+冻结首行），直接走预设更完整
            if (t.label === '商务蓝') return { preset: true, label: '商务蓝' };
            return { fill: t.fill, border: t.border, label: t.label };
        }
    }
    // 只说"主题/配色/风格"没点名 → 默认 palette（business_blue）即后端预设
    if (/(主题|配色|风格|色系)/.test(m)) return { preset: true, label: '商务蓝' };
    return null;
}

/**
 * 解析 excel_get_sheets 的返回值 → [{name, usedRange}]（仅可见表，后端已过滤隐藏表）。
 * 桥返回形态：{"success":true,"result":"[{name,index,usedRange},…]","isError":false}，
 * result 可能是 JSON 字符串也可能已被解析成数组 → 逐层下钻，两种形态都认。
 */
function _parseSheetList(raw) {
    let o = raw;
    for (let i = 0; i < 4; i++) {
        if (o == null) return [];
        if (typeof o === 'string') {
            const s = o.trim();
            if (!s || (s.charAt(0) !== '{' && s.charAt(0) !== '[')) return [];
            try { o = JSON.parse(s); } catch { return []; }
        }
        if (Array.isArray(o)) {
            return o.filter(x => x && typeof x.name === 'string' && x.name.trim())
                .map(x => ({
                    name: x.name.trim(),
                    usedRange: (typeof x.usedRange === 'string' ? x.usedRange : '').replace(/\$/g, '').trim()
                }));
        }
        if (typeof o !== 'object') return [];
        if (o.result != null) { o = o.result; continue; }
        if (Array.isArray(o.sheets)) { o = o.sheets; continue; }
        if (Array.isArray(o.content)) {
            const t = o.content.find(c => c && typeof c.text === 'string');
            if (t) { o = t.text; continue; }
        }
        return [];
    }
    return [];
}

// ── 选区范围判断与降级阶梯：统一委托 services/range-resolver.js ──
// 单一事实源（agent-loop 的无选区兜底、skill-spec-build 的选区优先判定共用）。
// 实现与踩坑背景见 range-resolver.js 头部注释（UsedRange 维度缓存对未激活表不刷新）。
// 模块缺失时全部落到"不可靠"（安全侧：不猜范围、不拿单格硬调），功能让位给正确性。
const _RR = () => ((window.App && App.ChatV2 && App.ChatV2.rangeResolver) || null);

function _colToNum(col) {
    const rr = _RR();
    return rr ? rr.colToNum(col) : 0;
}

function isReliableTableRange(addr) {
    const rr = _RR();
    return rr ? rr.isReliableTableRange(addr) : false;
}

function _normRangeAddr(addr) {
    const rr = _RR();
    return rr ? rr.normRangeAddr(addr) : String(addr || '').replace(/\$/g, '').trim();
}

function _readRangeAddr(raw) {
    const rr = _RR();
    return rr ? rr.readRangeAddr(raw) : null;
}

let _sheetRangeLadder = null;   // 懒加载：解析器只依赖 callLocalToolWithFallback，可缓存复用

/**
 * 【选区范围判断·降级阶梯】解析一张表的数据范围，逐级升级、每级都校验：
 *   ① seedRange（excel_get_sheets 的 usedRange）可靠 → 直接用；
 *   ② excel_get_used_range {sheet}（仍是 UsedRange，但对已激活表是准的）；
 *   ③ excel_read_range {sheet, range:'A1', autoProbe:true} —— 后端对单格锚点做
 *      CurrentRegion 连续区扩展，【不依赖工作表激活态】，是破解 UsedRange 缓存
 *      不刷新（拆分新表返回 G30 单格）的关键一跳；maxRows=1 只取地址不搬数据；
 *   ④ 全部不可靠 → 返回 null，由调用方【跳过该表并说明原因】（不猜、不格式错范围）。
 * @returns {Promise<{range:string, via:string}|null>}
 */
function resolveSheetRangeLadder(sheetName, seedRange) {
    const rr = _RR();
    if (!rr) return Promise.resolve(null);
    if (!_sheetRangeLadder) _sheetRangeLadder = rr.createSheetRangeLadder(callLocalToolWithFallback);
    return _sheetRangeLadder(sheetName, seedRange);
}

/**
 * 组装链式美化的 excel_format_range 参数：
 *   - format 必填（schema required）：detectFormatIntent 没给出具体样式 → '{}' 走后端商务蓝预设；
 *   - range 必填：调用方给不出（拆分产物表事先不知道已用范围）时留空 + autoRange=true，
 *     由执行期 get_used_range 补；⚠ 绝不能空 range 也不 autoRange 直调——后端会用"当前选区"
 *     兜底，把用户停在别的表上的选区格式化掉（silently 做错表）。
 */
function _buildChainFormatArgs(fmtArgs, sheet, range) {
    const a = {};
    const f = fmtArgs || {};
    if (f.format && typeof f.format === 'object') a.format = JSON.stringify(f.format);
    else if (typeof f.format === 'string' && f.format) a.format = f.format;
    else a.format = '{}';
    if (f.alignment && typeof f.alignment === 'object') a.alignment = JSON.stringify(f.alignment);
    else if (typeof f.alignment === 'string' && f.alignment) a.alignment = f.alignment;
    if (sheet) a.sheet = sheet;
    if (range) a.range = range;
    else a.autoRange = true;   // 执行期负责补真实 range；此处仅标记意图
    return a;
}

// ── 非 Excel 宿主（PPT/Word）"选中表格格式化"参数解析 ──
// 产出 ppt/word_format_selected_table_cell 的整表参数；未提及项留空、由后端按需处理。
// 关键：刻意不返回 row/col → 后端走"整表生效"分支，用户无需再报行列号。
function buildHostTableFormatArgs(message) {
    const m = String(message || '');
    const a = {};
    if (/居中|center/i.test(m)) a.align = 'center';
    else if (/左对齐|靠左|居左/i.test(m)) a.align = 'left';
    else if (/右对齐|靠右|居右/i.test(m)) a.align = 'right';
    else if (/对齐/.test(m)) a.align = 'center';   // 只说"对齐"→ 默认居中

    const removeFill = /消除|去掉|清除|移除|删除|不要底色|不要底纹|无底色|无底纹|取消底色|取消底纹|去除|去底色|去底纹|清底色|清底纹/.test(m);
    const wantFill = /颜色|底色|底纹|背景色|填充色|高亮|填充|灰底/.test(m) && !removeFill;
    if (wantFill) {
        const hex = m.match(/#([0-9a-fA-F]{6})/);
        const names = [['红', '#FF0000'], ['蓝', '#0000FF'], ['绿', '#00B050'], ['黄', '#FFFF00'], ['橙', '#FFA500'], ['紫', '#7030A0'], ['灰', '#808080'], ['黑', '#000000'], ['白', '#FFFFFF']];
        let fill = hex ? ('#' + hex[1]) : null;
        if (!fill) { for (const kv of names) { if (m.includes(kv[0])) { fill = kv[1]; break; } } }
        a.headerFillColor = fill || '#D9D9D9';       // 未指定具体色 → 默认浅灰表头
        a.headerBold = true;
    }
    if (removeFill) {
        a.clearFill = true;                          // 明确"消除/去掉底色"→ 宿主清除单元格填充
    }
    // ── 样式方案（版式 × 配色 × 密度）──
    // 旧实现只有 THEME_MAP（6 个主题，只换颜色）→ 换主题等于换配色、版式永远一样（用户反馈根因）。
    // 现在统一走 config/table-style-presets.js：10 个版式 × 8 套配色 × 3 档行高，并支持
    // 首列强调 / 末行合计 / 数字右对齐 / 垂直居中等开关。用户显式给的参数（颜色/对齐/字号…）优先于预设。
    const TSP = (typeof window !== 'undefined' && window.App && window.App.ChatV2 && window.App.ChatV2.tableStylePresets) || null;
    let presetArgs = null;
    if (TSP) {
        try {
            const det = TSP.detect(m);
            const hasSignal = !!(det.layoutId || det.paletteId || det.density ||
                det.zebra !== null || det.firstColBold !== null || det.lastRowEmphasis !== null ||
                det.bodyNumberRight !== null || det.headerMode || det.valign);
            if (hasSignal) presetArgs = TSP.toHostArgs(TSP.resolve(m));
        } catch { presetArgs = null; }
    }
    if (presetArgs) {
        Object.keys(presetArgs).forEach(k => { if (a[k] === undefined) a[k] = presetArgs[k]; });
    }
    // 普通"美化/设置表格格式/样式"但没显式给边框/没要求消除底纹、也没命中样式预设 → 默认补一道浅灰网格边框 + 隔行底纹，
    // 否则原表无边框时只会给表头灰底+加粗，正文仍是裸表格（用户反馈"太原始"）。
    if (/(美化|表格格式|样式)/.test(m) && !a.borderColor && !removeFill) {
        a.borderColor = '#BFBFBF';
        if (a.zebra === undefined) a.zebra = true;
    }
    // Word/PPT：保留样式预设给出的版式字段（三线表 / 行高 / 正文对齐 / 主题色）。
    // 仅剥离易触发合并格强退的逐格强调项；边框白名单含 threeLine。
    if (!/行高|行距/.test(m) && !(presetArgs && presetArgs.rowHeight != null)) delete a.rowHeight;
    if (!/垂直对齐|顶端对齐|垂直居中/.test(m) && !(presetArgs && presetArgs.valign)) delete a.valign;
    delete a.firstColBold;
    delete a.lastRowBold;
    if (!(presetArgs && presetArgs.bodyAlign)) delete a.bodyAlign;
    delete a.headerItalic;
    if (a.borderStyle && !/^(all|outer|none|threeLine)$/i.test(String(a.borderStyle))) a.borderStyle = 'all';
    // Word 宿主：美化默认补边框；勿覆盖已由预设给出的 threeLine 等
    try {
        const _hk = (typeof window !== 'undefined' && window.App && App.HostContext && App.HostContext.host)
            || (typeof window !== 'undefined' && window.SmartPilotHost && window.SmartPilotHost.kind)
            || (typeof URLSearchParams !== 'undefined' && new URLSearchParams(location.search).get('host'))
            || '';
        if (String(_hk).toLowerCase() === 'word' && /(美化|表格格式|样式)/.test(m)) {
            if (!a.borderStyle) a.borderStyle = 'all';
            if (!a.align && /居中|对齐/.test(m)) a.align = 'center';
        }
    } catch (_e) { /* ignore */ }
    // 文本突出颜色(高亮笔)：与底色/底纹是不同概念，单独成项。
    // "高亮"仍保留给底色填充（历史语义），此处只认"突出颜色/突出显示/高亮色/荧光"等明确措辞。
    if (/突出颜色|突出显示|高亮色|荧光|荧光笔|文本.*高亮|高亮.*文本/i.test(m)) {
        const hexH = m.match(/#([0-9a-fA-F]{6})/);
        const namesH = [['红', '#FF0000'], ['蓝', '#0000FF'], ['绿', '#00B050'], ['黄', '#FFFF00'], ['橙', '#FFA500'], ['紫', '#7030A0'], ['青', '#00B0B0'], ['粉', '#FF00FF'], ['黑', '#000000'], ['白', '#FFFFFF']];
        let hl = hexH ? ('#' + hexH[1]) : null;
        if (!hl) { for (const kv of namesH) { if (m.includes(kv[0])) { hl = kv[1]; break; } } }
        a.highlightColor = hl || '#FFFF00';          // 未指定具体色 → 默认黄（Word/PPT 高亮默认）
    }
    if (/加粗|加黑|表头|标题行|首行|美化|样式/.test(m)) a.headerBold = true;
    const fs = m.match(/字号\s*(\d{1,3})/);
    if (fs) a.fontSize = parseInt(fs[1], 10);
    const fonts = ['微软雅黑', '等线', '宋体', '黑体', '楷体', '仿宋'];
    for (const f of fonts) { if (m.includes(f)) { a.fontName = f; break; } }
    return a;
}

// ── 多要求格式化的"覆盖度"判定（防"只做了一部分就宣称完成"）──
// 确定性硬短路只能识别/覆盖 excel_format_range 的一个子集（bold/italic/border/wrapText/fontSize/
// fontName/fontColor/bgColor/alignment）。当用户一句话里同时提了多项要求时，硬短路会把"覆盖不到"
// 的要求静默丢弃——实测："美化…设置表头底色、边框、对齐方式与列宽"：
//   · 底色   → 紧跟"、"导致 buildFormatArgs 的取值正则取不到 → bgColor 丢失
//   · 对齐方式 → 不含"居中"，alignment 为空
//   · 列宽   → 后端 FormatRange 末尾已统一 Columns.AutoFit（不再依赖单独 excel_autofit）
// 只执行了边框，且直接 CallLocalTool 绕过 LLM → agent-loop 的多步能力与"要求覆盖度自检"全部用不上。
// 因此：一旦存在"用户提了、但确定性参数覆盖不到"的格式要求，就放弃硬短路，放行下游 agent-loop 多步处理。
// @param {string} message 用户原始消息
// @param {object} fmtArgs buildFormatArgs 的结果 { format?, alignment? }
// @returns {string[]} 未被覆盖的格式要求名（空数组 = 完全覆盖，可安全硬短路）
function uncoveredFormatAspects(message, fmtArgs) {
    const m_ = (message || '').toLowerCase();
    const fmt = (fmtArgs && fmtArgs.format) || {};
    const al = (fmtArgs && fmtArgs.alignment) || {};
    const miss = [];
    if (/加粗/.test(m_) && !fmt.bold) miss.push('加粗');
    if (/(倾斜|斜体)/.test(m_) && !fmt.italic) miss.push('斜体');
    if (/边框/.test(m_) && !fmt.borderStyle) miss.push('边框');
    if (/(自动换行|换行)/.test(m_) && !fmt.wrapText) miss.push('自动换行');
    if (/(字号|\d{1,3}\s*号字)/.test(m_) && !fmt.fontSize) miss.push('字号');
    if (/字体颜色/.test(m_) && !fmt.fontColor) miss.push('字体颜色');
    if (/(填充色|背景色|底色|底纹|填充)/.test(m_) && !fmt.bgColor) miss.push('底色');
    if (/(居中|对齐)/.test(m_) && !al.horizontal && !al.vertical) miss.push('对齐');
    // 列宽：走 format_range 时后端会 AutoFit，不再因「列宽」二字单独挡硬短路。
    // 行高自适应仍需 excel_autofit(type=row) 等，硬短路盖不住 → 交 agent-loop
    if (/行高/.test(m_) && fmt.rowHeight == null) miss.push('行高');
    return miss;
}

// 从选区范围字符串取首行范围（"把表头加粗"→选区第一行）。
// 支持 "$A$1:$G$201" / "A1:G201" / "Sheet1!A1:G201" / 单格 "A1"；统一返回带 $ 锚定首行范围。
// 仅用于"无显式区域但有选区"的确定性兜底——把选区首行（而非后端选区兜底误格式化的整表）作为目标。
function firstRowRangeOf(selRange) {
    if (!selRange || typeof selRange !== 'string') return null;
    const cleaned = selRange.replace(/^[^!]*!/, ''); // 去 sheet 前缀
    const cellRe = /^\$?([A-Za-z]+)\$?(\d+)$/;
    const parts = cleaned.split(':');
    if (parts.length === 1) {
        const c = parts[0].match(cellRe);
        return c ? `$${c[1]}$${c[2]}` : null;
    }
    const a = parts[0].match(cellRe);
    const b = parts[1].match(cellRe);
    if (!a || !b) return null;
    return `$${a[1]}$${a[2]}:$${b[1]}$${a[2]}`;
}

// ── 筛选条件解析（确定性硬短路辅助）──
// 从"筛选 订单金额>3000 且 区域=华东"这类消息里抽出 { 列名: 条件值 }。
// · 键保持用户原文（中文列名），由后端按表头解析成列（精确/包含），前端不猜列字母；
// · 值内嵌运算符：'=' 直接给值（等值），其余保留前缀（>3000 / <=100）交给 Excel AutoFilter。
function parseFilterCriteria(message) {
    const out = {};
    if (typeof message !== 'string' || !message) return out;
    // 列名：不含空格，≤12 字
    const NAME = '([\\u4e00-\\u9fa5A-Za-z_][\\u4e00-\\u9fa5A-Za-z0-9_]{0,11}?)';
    // 运算符：长的在前（== 先于 =，<> 先于 <），否则会被拆成 "=" + "=值"
    const OP = '(>=|<=|<>|!=|==|≥|≤|>|<|＝|=)';
    // 连接词：多条件分隔
    const CONN = '(?:且|并且|而且|以及|同时|并|和|、|，|,)';
    // 值：惰性匹配；只有遇到"分隔/结束"或"连接词+下一个 列名+运算符"才停。
    // 这样">3000且区域=华东"（无空格）也能正确切成两个条件，而"家电和数码"这种含连接词的取值不会被截断。
    const re = new RegExp(NAME + '\\s*' + OP + '\\s*([^\\s，。、；;,]+?)(?=$|[\\s。；;]|' + CONN + NAME + '\\s*' + OP + ')', 'g');
    let m;
    while ((m = re.exec(message)) !== null) {
        let name = String(m[1] || '').trim();
        let op = m[2];
        const val = String(m[3] || '').trim();
        // 去掉误吞进列名里的命令词/连接词（如"筛选订单金额"、"且区域"）
        name = name.replace(/^(筛选|过滤|筛出|筛选出|只看|仅显示|显示)/, '')
                   .replace(/^(且|并|并且|而且|以及|和|与|同时|、)+/, '')
                   .trim();
        if (!name || !val) continue;
        if (op === '≥') op = '>=';
        else if (op === '≤') op = '<=';
        else if (op === '＝' || op === '==') op = '=';
        out[name] = (op === '=') ? val : (op + val);
    }
    return out;
}

/**
 * 【可观测性】把 McpDispatcher 的关键判定写进宿主日志文件。
 * 背景：ChatLog 只写 WebView2 console，生产环境完全不可见，而"硬短路为什么没命中/为什么没执行"
 * 正是最难靠猜定位的一类问题。Excel 经 ChatLogger 写 chat_log_*.txt；
 * 共享宿主（PPT/Word）经 HostBridgeBase.LogMessage 写 host_bridge.log。
 */
function hostLog(level, message) {
    try {
        const b = (typeof App !== 'undefined' && App.excelBridge)
            || (typeof window !== 'undefined' && window.excelBridge)
            || (typeof window !== 'undefined' && window.chrome && window.chrome.webview
                && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
        if (b && typeof b.LogMessage === 'function') {
            const r = b.LogMessage(level || 'INFO', '[McpDispatcher] ' + message);
            if (r && typeof r.catch === 'function') r.catch(function () { /* 宿主日志失败可忽略 */ });
        }
    } catch (_e) { /* 宿主日志失败绝不影响主流程 */ }
}

/**
 * 收集当前可用的桥对象候选（去重），供 CallLocalTool 多来源回退调用。
 * 背景：某些环境下 window.BridgeAccessor.getBridge() 返回的包装桥对个别方法不可调用
 * （报 "e is not a function"），而裸宿主对象 App.excelBridge / window.excelBridge 可用
 * （hostLog 正是走裸桥且一直正常）。依次尝试可显著提升鲁棒性。
 */
function _collectBridgeCandidates() {
    const out = [];
    const push = function (b, label) {
        if (!b) return;
        for (let i = 0; i < out.length; i++) { if (out[i].b === b) return; }
        out.push({ b: b, label: label });
    };
    try { if (window.BridgeAccessor && typeof window.BridgeAccessor.getBridge === 'function') push(window.BridgeAccessor.getBridge(), 'BridgeAccessor.getBridge()'); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('MCP', '操作失败:', _ && _.message); }
    try { if (typeof App !== 'undefined') push(App.excelBridge, 'App.excelBridge'); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('MCP', '操作失败:', _ && _.message); }
    try { push(window.excelBridge, 'window.excelBridge'); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('MCP', '操作失败:', _ && _.message); }
    try { push(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge, 'hostObjects.excelBridge'); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('MCP', '操作失败:', _ && _.message); }
    return out;
}

/**
 * 调用宿主 CallLocalTool（多来源回退）。
 * @returns {Promise<{raw:any, src:string}>}
 * @throws 当所有可用来源都失败时抛最后一个错误（保留自然错误信息）。
 */
/**
 * 【合并结果瘦身·咽喉拦截 2026-10-09】递归找到"合并结果对象"（带 mergedRows/dataRows/sourceSheets
 * 标记）并剥离其全量 data 矩阵 → previewRows(前10行)。之所以放在 callLocalToolWithFallback：
 * dispatcher 硬短路（L760/801/960）与 agent-loop 兜底直调（L3589）全部经过本函数，
 * 在此单点拦截 = 两条路径全覆盖，下游任何漏截断都拿不到大矩阵，无从复述成裸 JSON。
 * 后端 SplitMerge.vb 已同步改为只回 previewRows；此处兜底兼容"旧 DLL 未重编"的过渡期。
 * @returns {boolean} 是否发生了剥离
 */
function _stripMergeDataMatrix(o, depth) {
    if (!o || typeof o !== 'object' || (depth || 0) > 4) return false;
    if (Array.isArray(o.data) &&
        (o.mergedRows !== undefined || o.dataRows !== undefined || Array.isArray(o.sourceSheets))) {
        if (o.previewRows === undefined) {
            o.previewRows = o.data.slice(0, 10);
            o.previewTruncated = o.data.length > o.previewRows.length;
        }
        delete o.data;
        return true;
    }
    const keys = ['result', 'data', 'content', 'outputs', 'payload'];
    for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        if (o[k] == null) continue;
        let inner = o[k];
        if (typeof inner === 'string') {
            const s = inner.trim();
            if (s.charAt(0) !== '{' && s.charAt(0) !== '[') continue;
            try { inner = JSON.parse(s); } catch (e) { continue; }
            if (_stripMergeDataMatrix(inner, (depth || 0) + 1)) { o[k] = JSON.stringify(inner); return true; }
            continue;
        }
        if (_stripMergeDataMatrix(inner, (depth || 0) + 1)) return true;
    }
    return false;
}

async function callLocalToolWithFallback(toolName, argsJson) {
    const cands = _collectBridgeCandidates();
    let lastErr = null;
    const tried = [];
    for (let i = 0; i < cands.length; i++) {
        const c = cands[i];
        try {
            if (!c.b || typeof c.b.CallLocalTool !== 'function') { tried.push(c.label + ':no-fn'); continue; }
            let raw = await c.b.CallLocalTool(toolName, argsJson);
            if (tried.length) hostLog('INFO', `[bridge-fallback] ${toolName} 经 ${c.label} 成功（先前失败: ${tried.join(',')}）`);
            // 【咽喉拦截】合并结果在跨桥边界立即瘦身，见 _stripMergeDataMatrix 注释
            if (toolName === 'excel_merge_sheets' && raw != null) {
                try {
                    if (typeof raw === 'string') {
                        if (raw.indexOf('"data"') >= 0) {
                            const o = JSON.parse(raw);
                            if (_stripMergeDataMatrix(o, 0)) raw = JSON.stringify(o);
                        }
                    } else if (typeof raw === 'object' && _stripMergeDataMatrix(raw, 0)) {
                        hostLog('INFO', '[bridge-fallback] excel_merge_sheets 结果已剥离全量 data 矩阵（→ previewRows 10 行）');
                    }
                } catch (_smErr) { /* 解析失败原样返回，下游守卫兜底 */ }
            }
            return { raw: raw, src: c.label };
        } catch (e) {
            lastErr = e;
            tried.push(c.label + ':' + ((e && e.message) || e));
        }
    }
    if (tried.length) { try { hostLog('WARN', `[bridge-fallback] ${toolName} 所有来源均失败: ${tried.join(' | ')}`); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('MCP', '操作失败:', _ && _.message); } }
    if (lastErr) throw lastErr;
    throw new Error('无可用桥（CallLocalTool 在所有来源上均不可用）');
}

// 统一硬短路执行结果的返回结构（单一标准）。
// 三条执行路径历史上各自返回不同 execResult 形态：
//   · 关键词硬短路（CallLocalTool）      → { ok, result, error, steps }
//   · MatchCapability 命中 → CallLocalTool → { ok, result, raw, elapsedMs }
//   · 降级 invokeTracked                → { ok, result, error, durationMs, toolCallId, elapsedMs }
// 调用方（orchestrator）虽已用 ?. / || 兜底容错，但契约不明确，新增路径易漏字段。
// 本助手把三者收敛为 {ok,result,error,raw,steps,elapsedMs} 六字段标准，缺失项给默认，
// 后续新增路径也走本助手，避免再次漂移。
function _parseMaybe(v) {
    if (typeof v !== 'string') return v;
    const s = v.trim();
    if (!s || (s.charAt(0) !== '{' && s.charAt(0) !== '[')) return v;
    try { return JSON.parse(s); } catch { return v; }
}

function _readDedupeBiz(root, depth) {
    if (root == null || depth > 5) return null;
    const node = _parseMaybe(root);
    if (!node || typeof node !== 'object' || Array.isArray(node)) return null;
    if (Array.isArray(node.duplicateRows) || node.deduplicated === true || node.removedDuplicateRows != null) return node;
    if (node.result != null) {
        const inner = _readDedupeBiz(node.result, depth + 1);
        if (inner) return inner;
    }
    if (Array.isArray(node.content)) {
        for (const c of node.content) {
            if (c && typeof c.text === 'string') {
                const inner = _readDedupeBiz(c.text, depth + 1);
                if (inner) return inner;
            }
        }
    }
    return null;
}

function _stampDedupe(root, patch, depth) {
    if (root == null || depth > 5) return null;
    if (typeof root !== 'object' || Array.isArray(root)) return null;
    const isBiz = Array.isArray(root.duplicateRows) || root.deduplicated === true || root.removedDuplicateRows != null;
    if (isBiz) {
        Object.assign(root, patch);
        return root;
    }
    if (typeof root.result === 'string') {
        const node = _parseMaybe(root.result);
        if (node && typeof node === 'object' && !Array.isArray(node)) {
            const hit = _stampDedupe(node, patch, depth + 1);
            if (hit) { root.result = JSON.stringify(node); return hit; }
        }
    } else if (root.result && typeof root.result === 'object' && !Array.isArray(root.result)) {
        const hit = _stampDedupe(root.result, patch, depth + 1);
        if (hit) return hit;
    }
    if (Array.isArray(root.content)) {
        for (const c of root.content) {
            if (c && typeof c.text === 'string') {
                const node = _parseMaybe(c.text);
                if (node && typeof node === 'object' && !Array.isArray(node)) {
                    const hit = _stampDedupe(node, patch, depth + 1);
                    if (hit) { c.text = JSON.stringify(node); return hit; }
                }
            }
        }
    }
    return null;
}

function _spanFromRange(addr) {
    const s = String(addr || '').replace(/\$/g, '');
    const m = s.match(/([A-Za-z]{1,3})\d+\s*:\s*([A-Za-z]{1,3})\d+/);
    if (!m) return null;
    return [m[1].toUpperCase(), m[2].toUpperCase()];
}

function _colLetter(n) {
    let s = '';
    let x = n;
    while (x > 0) {
        const r = (x - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        x = Math.floor((x - 1) / 26);
    }
    return s || 'A';
}

/** 用去重结果里的绝对行号把重复行标黄。不删行。 */
async function _highlightDuplicateRows(parsed, hardArgs) {
    const biz = _readDedupeBiz(parsed, 0);
    const rows = (biz && Array.isArray(biz.duplicateRows) ? biz.duplicateRows : [])
        .map(function (n) { return Number(n); })
        .filter(function (n) { return Number.isFinite(n) && n >= 1; });
    const uniq = Array.from(new Set(rows)).sort(function (a, b) { return a - b; });
    let cols = _spanFromRange(biz && biz.range) || _spanFromRange(hardArgs && hardArgs.range);
    if (!cols) {
        const n = (biz && Array.isArray(biz.headers) && biz.headers.length) || 1;
        cols = ['A', _colLetter(n)];
    }
    const sheet = (biz && (biz.worksheet || biz.sheet)) || (hardArgs && hardArgs.sheet) || '';
    const groups = [];
    uniq.forEach(function (r) {
        const last = groups[groups.length - 1];
        if (last && r === last[1] + 1) last[1] = r;
        else groups.push([r, r]);
    });
    let failed = 0;
    for (let i = 0; i < groups.length; i++) {
        const a = groups[i][0];
        const b = groups[i][1];
        const fmt = { range: cols[0] + a + ':' + cols[1] + b, format: JSON.stringify({ bgColor: '#FFFF00', bgScope: 'all' }) };
        if (sheet) fmt.sheet = sheet;
        try {
            await callLocalToolWithFallback('excel_format_range', JSON.stringify(fmt));
        } catch (e) {
            failed++;
            ChatLog.warn?.('McpDispatcher', '标黄重复行失败: ' + ((e && e.message) || e));
        }
    }
    let writtenTo = '（已标黄重复行，未删除数据）';
    if (!uniq.length) writtenTo = '（未发现重复行，未删除数据）';
    else if (failed === groups.length) writtenTo = '（标黄失败，未删除数据）';
    else if (failed) writtenTo = '（部分重复行标黄失败，未删除数据）';
    _stampDedupe(parsed, { markOnly: true, highlightedRows: uniq, writtenTo: writtenTo }, 0);
}

/** 长的色名放前面，避免「浅蓝」被「蓝」截走。 */
const MARK_FILL_COLORS = [
    ['浅红', 'FF9999', '浅红色'], ['深红', 'C00000', '深红色'],
    ['浅蓝', 'BDD7EE', '浅蓝色'], ['深蓝', '1F4E79', '深蓝色'],
    ['浅绿', 'C6EFCE', '浅绿色'], ['深绿', '006600', '深绿色'],
    ['红', 'FF0000', '红色'], ['橙', 'FFC000', '橙色'], ['黄', 'FFFF00', '黄色'],
    ['绿', '00B050', '绿色'], ['青', '00B0F0', '青色'], ['蓝', '5B9BD5', '蓝色'],
    ['紫', '7030A0', '紫色'], ['粉', 'FFC0CB', '粉色'], ['灰', 'BFBFBF', '灰色'],
    ['黑', '000000', '黑色'], ['白', 'FFFFFF', '白色'], ['棕', 'C65911', '棕色'],
    ['金', 'FFD966', '金色']
];
const MARK_COLOR_WORD = '(?:深|浅|亮)?(?:红|橙|黄|绿|青|蓝|紫|粉|灰|黑|白|棕|金)色?';

/** 「标记王五」「将王五标成绿色」「换成蓝色」：抽出要上色的字面值。重复行另走去重高亮。 */
function extractMarkLiteral(message) {
    const raw = String(message || '').trim();
    if (!raw || /重复/.test(raw)) return null;
    // 主题类请求排除："将工作簿中的所有工作表转为商务蓝"曾被本函数误判成"把'工作簿中的
    // 所有工作表转为商务'标记成蓝色"（quoted 贪婪吞掉动词、结尾的"蓝"命中颜色词），
    // 整条请求被 excel_apply_conditional_format 截走（实测 2026-10-22）。
    if (/(删除|移除|去掉|清洗|清理|去重|统计|查找|搜索|筛选|排序|主题|配色|风格|色系|商务蓝|科技绿|活力橙|优雅紫|极简灰|金融红|墨黑金|晴空青|排版)/.test(raw)) return null;
    const strip = function (s) {
        let v = String(s || '').trim().replace(/^[「『"'“]+|[」』"'”]+$/g, '');
        v = v.replace(/(所在的?|的)?(单元格|格子|这一列|这列|这行|记录|数据|内容)$/g, '').trim();
        if (!v || v.length > 30 || /^(一下|出来|这些|这个|全部|所有|重复行|重复)$/.test(v)) return null;
        if (new RegExp('^' + MARK_COLOR_WORD + '$').test(v)) return null;
        return v;
    };
    const quoted = '[「『"\'“]?([^」』"\'”\\s，。,.]{1,30})[」』"\'”]?';
    const colorAct = '(?:标|涂|高亮|标记|标出|标注|改成|换成|设为|设成)(?:成|为|上)?' + MARK_COLOR_WORD;
    let m = raw.match(new RegExp('^(?:请|帮我|帮忙)?(?:把|将)\\s*' + quoted + '\\s*(?:' + colorAct + ')'));
    if (m) return strip(m[1]);
    m = raw.match(new RegExp('^(?:请|帮我|帮忙)?\\s*(?:标|涂)(?:成|为|上)?' + MARK_COLOR_WORD + '\\s*' + quoted));
    if (m) return strip(m[1]);
    m = raw.match(/^(?:请|帮我|帮忙)?\s*(?:标记|高亮|标出|标注)(?:一下|出来|出)?\s*[「『"'“]?([^」』"'”\s，。！？,.]{1,30})/);
    if (m) return strip(m[1]);
    // 「将销售一部黄色」「把华东区填成绿色」：把/将 + 值 + 颜色，中间可以不写标/涂。
    m = raw.match(new RegExp('^(?:请|帮我|帮忙)?(?:把|将)\\s*' + quoted
        + '(?:的)?(?:单元格|格子|所在行)?\\s*(?:填(?:充)?|涂|标|高亮|标记|改|换|设)?(?:成|为|上)?\\s*(?:'
        + MARK_COLOR_WORD + ')\\s*[。！？]?$'));
    if (m) return strip(m[1]);
    return null;
}

/** 「移除销售一部数据」「把王五删掉」：要整行删掉的单元格字面值。重复行、空行、第 N 行、删列不走这里。 */
function extractRemoveRowValue(message) {
    const raw = String(message || '').trim().replace(/[。！？!?\s]+$/g, '');
    if (!raw || raw.length > 40) return null;
    if (/重复|空行|空白|清洗|清理|排序|筛选|图表|统计|查找|搜索|标记|高亮|标红|标黄/.test(raw)) return null;
    if (/第\s*\d+\s*行|[A-Za-z]{1,3}\s*列|整列|整行|全部|所有数据|整个表|整张表/.test(raw)) return null;
    const strip = function (s) {
        let v = String(s || '').trim().replace(/^[「『"'“]+|[」』"'”]+$/g, '');
        v = v.replace(/(?:的)?数据(?:从|于)?(?:单元格|格子)(?:中|里|内)?$/g, '');
        v = v.replace(/(?:从|于)?(?:所在的?)?(?:单元格|格子)(?:中|里|内)?$/g, '');
        v = v.replace(/(?:的)?(?:数据|记录|内容|这些|那几行|这几行|行)$/g, '').trim();
        if (!v || v.length > 30 || /^(一下|这些|这个|数据|记录|行|内容|全部|单元格)$/.test(v)) return null;
        return v;
    };
    let m = raw.match(/^(?:请|帮我|帮忙)?(?:移除|删除|去掉|清除|删掉)\s*[「『"'“]?([^」』"'”\s，,。]{1,30})/);
    if (m) return strip(m[1]);
    // 「将销售一部数据从单元格中移除」：值在「数据 / 从单元格」前面，不能把后半句收进要匹配的字。
    m = raw.match(/^(?:请|帮我|帮忙)?(?:把|将)\s*[「『"'“]?(.+?)[」』"'”]?\s*(?:的)?(?:数据|记录|内容)?\s*(?:从|于)?(?:所在的?)?(?:单元格|格子)?(?:中|里|内)?\s*(?:移除|删除|去掉|清除|删掉)\s*$/);
    if (m) return strip(m[1]);
    return null;
}

function _rangeStartRow(addr) {
    const m = String(addr || '').replace(/\$/g, '').match(/[A-Za-z]+(\d+)/);
    return m ? parseInt(m[1], 10) : 1;
}

function _digReadTable(raw) {
    let o = raw;
    for (let i = 0; i < 6; i++) {
        o = _parseMaybe(o);
        if (o == null) return null;
        if (typeof o !== 'object') return null;
        if (Array.isArray(o.result) && (typeof o.range === 'string' || Array.isArray(o.headers))) return o;
        if (Array.isArray(o.content)) {
            const t = o.content.find(function (c) { return c && typeof c.text === 'string'; });
            if (t) { o = t.text; continue; }
        }
        if (o.result != null && !Array.isArray(o.result)) { o = o.result; continue; }
        return null;
    }
    return null;
}

function _rowGroupsFromBottom(rows) {
    const sorted = rows.slice().sort(function (a, b) { return a - b; });
    const groups = [];
    let i = 0;
    while (i < sorted.length) {
        let j = i;
        while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
        groups.push({ row: sorted[i], count: j - i + 1 });
        i = j + 1;
    }
    groups.reverse();
    return groups;
}

/** 读已用区域，删掉任一格等于该值的整行。从最后一行往上删，避免行号错位。 */
async function _deleteRowsByValue(value, sheet) {
    const readArgs = { hasHeaders: false, skipEmptyRows: false, returnFormat: 'array', maxRows: 10000 };
    if (sheet) readArgs.sheet = sheet;
    const got = await callLocalToolWithFallback('excel_read_range', JSON.stringify(readArgs));
    const table = _digReadTable(got && got.raw);
    if (!table || !Array.isArray(table.result)) {
        return { success: false, removedValue: value, error: '没有读到工作表数据，行没有删除' };
    }
    const start = _rangeStartRow(table.range);
    const hit = [];
    table.result.forEach(function (row, idx) {
        const cells = Array.isArray(row) ? row : (row && typeof row === 'object' ? Object.keys(row).map(function (k) { return row[k]; }) : []);
        const matched = cells.some(function (cell) { return cell != null && String(cell).trim() === value; });
        if (matched) hit.push(start + idx);
    });
    if (!hit.length) {
        return {
            success: true,
            removedValue: value,
            removedRows: [],
            removedCount: 0,
            sheet: table.worksheet || sheet || '',
            range: table.range || ''
        };
    }
    const groups = _rowGroupsFromBottom(hit);
    for (let g = 0; g < groups.length; g++) {
        const del = { row: groups[g].row, count: groups[g].count };
        if (sheet || table.worksheet) del.sheet = sheet || table.worksheet;
        await callLocalToolWithFallback('excel_delete_row', JSON.stringify(del));
    }
    return {
        success: true,
        removedValue: value,
        removedRows: hit,
        removedCount: hit.length,
        sheet: table.worksheet || sheet || '',
        range: table.range || '',
        writtenTo: '已直接修改'
    };
}

function markFillColor(message) {
    const t = String(message || '');
    for (let i = 0; i < MARK_FILL_COLORS.length; i++) {
        if (t.indexOf(MARK_FILL_COLORS[i][0]) >= 0) {
            return { hex: MARK_FILL_COLORS[i][1], name: MARK_FILL_COLORS[i][2] };
        }
    }
    return { hex: 'FFFF00', name: '黄色' };
}

/** 从 excel_get_used_range 返回取地址（实现见 services/range-resolver.js） */
function _usedRangeAddr(raw) {
    const rr = _RR();
    return rr ? rr.usedRangeAddr(raw) : null;
}

function _schemaBrief(toolRegistry, toolName) {
    const tool = toolRegistry && typeof toolRegistry.get === 'function' ? toolRegistry.get(toolName) : null;
    if (!tool) return '';
    const schema = tool.inputSchema || tool.parameters;
    const lines = [];
    if (schema && schema.properties && typeof schema.properties === 'object') {
        const req = new Set(Array.isArray(schema.required) ? schema.required : []);
        Object.keys(schema.properties).forEach(function (k) {
            const v = schema.properties[k] || {};
            const desc = String(v.description || v.type || '').replace(/\s+/g, ' ').slice(0, 160);
            lines.push('- ' + k + (req.has(k) ? '（必填）' : '') + (desc ? '：' + desc : ''));
        });
    } else if (Array.isArray(schema)) {
        schema.forEach(function (p) {
            if (!p || !p.name) return;
            const desc = String(p.description || p.type || '').replace(/\s+/g, ' ').slice(0, 160);
            lines.push('- ' + p.name + (p.required ? '（必填）' : '') + (desc ? '：' + desc : ''));
        });
    }
    const about = tool.description ? String(tool.description).replace(/\s+/g, ' ').slice(0, 220) : '';
    return (about ? '工具说明：' + about + '\n' : '') + (lines.length ? '参数：\n' + lines.join('\n') : '');
}

function _asPlainObject(v) {
    if (v && typeof v === 'object' && !Array.isArray(v)) return v;
    if (typeof v === 'string') {
        const s = v.trim();
        if (s.charAt(0) === '{') {
            try {
                const o = JSON.parse(s);
                if (o && typeof o === 'object' && !Array.isArray(o)) return o;
            } catch (_) { /* 不是 JSON 对象 */ }
        }
    }
    return null;
}

function _hex6(v) {
    if (typeof v !== 'string') return null;
    const m = v.trim().match(/^#?([0-9a-fA-F]{6})$/);
    return m ? m[1].toUpperCase() : null;
}

/** 模型照抄 schema 示例（列 A、范围 A1、默认黄）且原话里没有这些词时，丢掉该值。 */
function _isSchemaPlaceholder(key, value, message) {
    const msg = String(message || '');
    const s = typeof value === 'string' ? value.trim() : '';
    if ((key === 'column' || key === 'dimColumn' || key === 'valueColumn' || key === 'dateColumn' || key === 'aggColumn')
        && /^[A-Za-z]{1,3}$/.test(s) && !new RegExp('(?:^|[^A-Za-z])' + s + '(?:$|[^A-Za-z])', 'i').test(msg)) {
        return true;
    }
    if (key === 'range' && /^\$?A\$?1(?::\$?[A-Za-z]{1,3}\$?\d+)?$/i.test(s) && !/A1/i.test(msg)) return true;
    const hex = _hex6(s);
    if ((key === 'color' || key === 'bgColor' || key === 'fontColor') && hex === 'FFFF00' && !/黄/.test(msg) && /红|橙|绿|青|蓝|紫|粉|灰|黑|白|棕|金/.test(msg)) {
        return true;
    }
    return false;
}

/**
 * 模型给出的参数覆盖种子值；模型没给的字段保留种子（无模型时行为不变）。
 * options / criteria 按字段合并。颜色只接受 6 位十六进制，中文色名留给模型转成色值。
 */
function mergeSemanticArgs(seed, ai, message) {
    const base = (seed && typeof seed === 'object') ? Object.assign({}, seed) : {};
    if (!ai || typeof ai !== 'object' || Array.isArray(ai)) return base;
    const lock = { headers: 1, rows: 1, filePath: 1, returnFormat: 1 };
    Object.keys(ai).forEach(function (k) {
        if (lock[k] && base[k] != null && base[k] !== '') return;
        let v = ai[k];
        if (v == null || v === '') return;
        if (typeof v === 'string' && (v === 'true' || v === 'false') && typeof base[k] === 'boolean') v = v === 'true';
        if (k === 'options' || k === 'criteria' || k === 'format' || k === 'alignment') {
            const fromAi = _asPlainObject(v);
            if (!fromAi) return;
            const fromSeed = _asPlainObject(base[k]) || {};
            const merged = Object.assign({}, fromSeed);
            Object.keys(fromAi).forEach(function (ik) {
                const iv = fromAi[ik];
                if (iv == null || iv === '') return;
                if ((ik === 'color' || ik === 'bgColor' || ik === 'fontColor')) {
                    const hex = _hex6(String(iv));
                    if (!hex || _isSchemaPlaceholder(ik, hex, message)) return;
                    merged[ik] = hex;
                    return;
                }
                if (_isSchemaPlaceholder(ik, iv, message)) return;
                merged[ik] = iv;
            });
            base[k] = JSON.stringify(merged);
            if (merged.value != null && merged.value !== '') base.markLiteral = String(merged.value);
            if (typeof ai.markColorName === 'string' && ai.markColorName.trim()) base.markColorName = ai.markColorName.trim();
            else if (merged.color && _hex6(String(_asPlainObject(seed && seed[k]) && _asPlainObject(seed[k]).color)) !== merged.color) {
                base.markColorName = '#' + merged.color;
            }
            return;
        }
        if (k === 'sortBy') {
            let arr = null;
            if (typeof v === 'string') {
                try { arr = JSON.parse(v); } catch (_) { arr = null; }
            } else if (Array.isArray(v)) arr = v;
            const col = arr && arr[0] && arr[0].column;
            if (col != null && _isSchemaPlaceholder('column', String(col), message)) return;
            base[k] = typeof v === 'string' ? v : JSON.stringify(v);
            return;
        }
        if (_isSchemaPlaceholder(k, v, message)) return;
        if (Array.isArray(v) || (v && typeof v === 'object')) base[k] = JSON.stringify(v);
        else base[k] = v;
    });
    return base;
}

function _contextHint(turn) {
    const ctx = (turn && turn.context) || {};
    const sel = ctx.selection || {};
    const headers = Array.isArray(sel.headers) ? sel.headers.slice(0, 24) : [];
    const sheet = sel.sheetName || (ctx.activeSheet && ctx.activeSheet.sheet) || '';
    const bits = [];
    if (sheet) bits.push('当前工作表：' + sheet);
    if (headers.length) bits.push('表头：' + headers.join('、'));
    return bits.join('\n');
}

/**
 * 带 reason 的 abort。不传 reason 时 signal.reason 是默认 DOMException，
 * message 就是一句没有信息量的 "signal is aborted without reason"，
 * 下游 catch 只能看到它，分不清是用户取消还是超时。
 */
function abortWithReason(ac, msg) {
    try { ac.abort(new DOMException(msg, 'AbortError')); }
    catch (_) { try { ac.abort(); } catch (_2) { /* ignore */ } }
}

/**
 * 工具已经选定之后，用模型按用户原话填写参数。
 * 没有可用模型、超时或返回不是 JSON 时，原样返回种子参数。
 */
async function refineArgsFromSemantics(toolName, message, seedArgs, turn, signal, toolRegistry) {
    const seed = (seedArgs && typeof seedArgs === 'object') ? seedArgs : {};
    try {
        const llm = (typeof window !== 'undefined' && window.App && App.ChatV2 && App.ChatV2.llmClient) || null;
        const config = (typeof window !== 'undefined' && window.App && App.config) || null;
        if (!llm || typeof llm.completeNonStream !== 'function' || !config) return seed;
        if (!config.apiUrl && !config.apiType && !config.model) return seed;
        const brief = _schemaBrief(toolRegistry, toolName);
        const ctx = _contextHint(turn);
        const sys = '你只输出一个 JSON 对象，作为工具参数。值必须来自用户原话，不要照抄工具说明里的示例。'
            + '用户没说清的字段不要输出。颜色输出 6 位十六进制（不要中文色名、不要默认黄色）。'
            + '列名用用户说的表头文字，不要改成 A。'
            + '条件格式 options 用对象，operator 取 equal、greaterThan、lessThan、between、notEqual。'
            + '用户要清洗、去重、移除且没说只预览时 writeBack 为 true、dryRun 为 false；只标记、不删除时 markOnly 为 true。';
        const user = '工具：' + toolName
            + (brief ? '\n' + brief : '')
            + (ctx ? '\n' + ctx : '')
            + '\n用户原话：' + String(message || '').slice(0, 500)
            + '\n已解析参数（只作参考，与原话冲突时以原话为准）：' + JSON.stringify(seed).slice(0, 800);
        const ac = new AbortController();
        const timer = setTimeout(function () { abortWithReason(ac, '语义补参超过 8 秒，改用原始参数'); }, 8000);
        if (signal) {
            if (signal.aborted) abortWithReason(ac, '本轮对话已取消');
            else if (typeof signal.addEventListener === 'function') signal.addEventListener('abort', function () { abortWithReason(ac, '本轮对话已取消'); }, { once: true });
        }
        let content = '';
        try {
            const evt = await llm.completeNonStream({
                messages: [{ role: 'system', content: sys }, { role: 'user', content: user }],
                config: config,
                model: config.model || '',
                signal: ac.signal,
                nativeTools: false
            });
            content = (evt && evt.content) || '';
        } finally {
            clearTimeout(timer);
        }
        let parsed = null;
        const ju = window.App && App.ChatV2 && App.ChatV2.jsonUtils;
        if (ju && typeof ju.extractJson === 'function') parsed = ju.extractJson(content);
        if (!parsed) {
            const s = String(content || '');
            const a = s.indexOf('{');
            const b = s.lastIndexOf('}');
            if (a >= 0 && b > a) {
                try { parsed = JSON.parse(s.slice(a, b + 1)); } catch (_) { parsed = null; }
            }
        }
        if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return seed;
        const merged = mergeSemanticArgs(seed, parsed, message);
        ChatLog.info?.('McpDispatcher', '语义补参 ' + toolName + ' ' + JSON.stringify(merged).slice(0, 300));
        return merged;
    } catch (e) {
        // 取消/超时是正常分支（用户打断或 8s 超时），不是故障：降级为 info，别刷 warn
        var _aborted = !!(e && (e.name === 'AbortError' || /signal is aborted/i.test(e.message || '')));
        var _msg = '语义补参跳过 ' + toolName + ': ' + ((e && e.message) || e);
        if (_aborted) { if (ChatLog.info) ChatLog.info('McpDispatcher', _msg); }
        else if (ChatLog.warn) ChatLog.warn('McpDispatcher', _msg);
        return seed;
    }
}

const QUARTER_MONTHS = { 1: [1, 3], 2: [4, 6], 3: [7, 9], 4: [10, 12] };
const QUARTER_DIGIT = { 一: 1, 二: 2, 三: 3, 四: 4, '1': 1, '2': 2, '3': 3, '4': 4 };

/** Q3 / 第三季度 / 2026年Q3 → 公历季度。没写年份时不限年，只限月份。 */
function parseQuarterPeriod(message) {
    const s = String(message || '');
    const re = /(20\d{2})\s*年?\s*(?:Q\s*([1-4])|第?\s*([一二三四1-4])\s*季度)|(?:Q\s*([1-4])|第?\s*([一二三四1-4])\s*季度)/i;
    const m = s.match(re);
    if (!m) return null;
    const token = m[2] || m[3] || m[4] || m[5];
    const quarter = QUARTER_DIGIT[token] || parseInt(token, 10);
    const span = QUARTER_MONTHS[quarter];
    if (!span) return null;
    const year = m[1] ? parseInt(m[1], 10) : null;
    return {
        year: year,
        quarter: quarter,
        startMonth: span[0],
        endMonth: span[1],
        label: (year ? year + '年' : '') + 'Q' + quarter
    };
}

function parseCellDate(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number' && v > 20000 && v < 80000) {
        const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
        return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, day: d.getUTCDate() };
    }
    const s = String(v).trim();
    let hit = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月(?:\s*(\d{1,2})\s*日)?/);
    if (hit) return { y: parseInt(hit[1], 10), m: parseInt(hit[2], 10), day: hit[3] ? parseInt(hit[3], 10) : 1 };
    hit = s.match(/(\d{4})[-/.](\d{1,2})(?:[-/.](\d{1,2}))?/);
    if (hit) return { y: parseInt(hit[1], 10), m: parseInt(hit[2], 10), day: hit[3] ? parseInt(hit[3], 10) : 1 };
    return null;
}

function _cellOf(row, header, index) {
    if (Array.isArray(row)) return row[index];
    if (row && typeof row === 'object') return row[header];
    return null;
}

function filterRowsByQuarter(headers, rows, period) {
    const hs = Array.isArray(headers) ? headers : [];
    const rs = Array.isArray(rows) ? rows : [];
    if (!period || !hs.length || !rs.length) return { rows: rs, applied: false, kept: rs.length, dateColumn: '' };
    let col = hs.findIndex(function (h) { return /日期|时间|月份|date|time/i.test(String(h || '')); });
    if (col < 0) {
        for (let c = 0; c < hs.length; c++) {
            let hits = 0;
            for (let r = 0; r < Math.min(rs.length, 8); r++) {
                if (parseCellDate(_cellOf(rs[r], hs[c], c))) hits++;
            }
            if (hits >= 2) { col = c; break; }
        }
    }
    if (col < 0) return { rows: rs, applied: false, kept: rs.length, dateColumn: '' };
    const name = hs[col];
    const kept = rs.filter(function (row) {
        const d = parseCellDate(_cellOf(row, name, col));
        if (!d) return false;
        if (period.year && d.y !== period.year) return false;
        return d.m >= period.startMonth && d.m <= period.endMonth;
    });
    if (!kept.length) return { rows: rs, applied: false, kept: 0, dateColumn: String(name), empty: true };
    return { rows: kept, applied: true, kept: kept.length, dateColumn: String(name) };
}

const TIME_GROUP_WORD = /^(年|年度|季度|季|月|月份|月度|周|星期|日|天|日期|时间)$/;

function findDateColumn(headers, rows) {
    const hs = Array.isArray(headers) ? headers : [];
    const rs = Array.isArray(rows) ? rows : [];
    let col = hs.findIndex(function (h) { return /日期|时间|月份|date|time/i.test(String(h || '')); });
    if (col < 0) {
        for (let c = 0; c < hs.length; c++) {
            let hits = 0;
            for (let r = 0; r < Math.min(rs.length, 8); r++) {
                if (parseCellDate(_cellOf(rs[r], hs[c], c))) hits++;
            }
            if (hits >= 2) { col = c; break; }
        }
    }
    return col >= 0 ? String(hs[col]) : '';
}

function _num(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (v == null || v === '') return null;
    const n = Number(String(v).replace(/,/g, '').replace(/%$/, ''));
    return isFinite(n) ? n : null;
}

function _aggFunc(message) {
    const s = String(message || '');
    if (/平均|均值/.test(s)) return 'avg';
    if (/计数|个数|条数|多少条/.test(s)) return 'count';
    if (/最大|最高/.test(s)) return 'max';
    if (/最小|最低/.test(s)) return 'min';
    return 'sum';
}

function _reduce(vals, func) {
    const nums = vals.map(_num).filter(function (n) { return n != null; });
    if (func === 'count') return vals.filter(function (v) { return v != null && v !== ''; }).length;
    if (!nums.length) return 0;
    if (func === 'avg') return Math.round(nums.reduce(function (a, b) { return a + b; }, 0) / nums.length * 100) / 100;
    if (func === 'max') return Math.max.apply(null, nums);
    if (func === 'min') return Math.min.apply(null, nums);
    return Math.round(nums.reduce(function (a, b) { return a + b; }, 0) * 100) / 100;
}

function _chartType(message) {
    const s = String(message || '');
    if (/圆环|环形/.test(s)) return 'xlDoughnut';
    if (/饼|占比/.test(s)) return 'xlPie';
    if (/折线|趋势/.test(s)) return 'xlLineMarkers';
    if (/条形/.test(s)) return 'xlBarClustered';
    if (/柱/.test(s)) return 'xlColumnClustered';
    return 'auto';
}

function _followOf(message) {
    const s = String(message || '');
    if (/ppt|pptx|幻灯片|演示文稿|汇报/i.test(s)) return 'ppt';
    if (/图表|柱状|柱形|饼图|折线|条形图|画图|做图|出图/.test(s)) return 'chart';
    if (/写入|写到|落到|存到|保存/.test(s)) return 'write';
    return '';
}

/**
 * 分组计划和后面要做的功能分开。
 * Q3 只缩小到 7–9 月；「按季度 / 按部门分组」才把行收成一组。
 */
function parseGroupPlan(message, headers) {
    const s = String(message || '');
    const hs = Array.isArray(headers) ? headers.map(function (h) { return String(h || ''); }) : [];
    const period = parseQuarterPeriod(s);
    let grain = '';
    if (/按\s*年|每年|各年|年度汇总|按年度/.test(s)) grain = 'year';
    else if (/按\s*季度|每季度|各季度|按季(?!度报)/.test(s)) grain = 'quarter';
    else if (/按\s*月|每月|各月|月度/.test(s)) grain = 'month';
    else if (/按\s*周|每周|各周/.test(s)) grain = 'week';
    else if (/按\s*日|每日|按天|各日/.test(s)) grain = 'day';
    let dim = '';
    const dm = s.match(/(?:按|以)\s*([^\s，。,、；;]{1,12}?)\s*(?:分组|汇总|统计|合计|分类)/);
    if (dm && dm[1] && !TIME_GROUP_WORD.test(dm[1])) {
        const token = dm[1];
        const hit = hs.filter(function (h) { return h && (h === token || h.indexOf(token) >= 0 || token.indexOf(h) >= 0); })
            .sort(function (a, b) { return b.length - a.length; })[0];
        dim = hit || token;
    }
    if (!dim) {
        const em = s.match(/各(部门|地区|区域|产品|城市|省份|渠道|客户|销售部)/);
        if (em) {
            const hit = hs.filter(function (h) { return h && h.indexOf(em[1]) >= 0; })[0];
            dim = hit || em[1];
        }
    }
    const valueHit = hs.filter(function (h) {
        return h && h !== dim && s.indexOf(h) >= 0 && !/日期|时间|date|time/i.test(h);
    }).sort(function (a, b) { return b.length - a.length; })[0] || '';
    const collapse = !!(grain || dim || /分组|分类汇总/.test(s));
    if (!period && !collapse) return null;
    const follow = _followOf(s);
    const label = dim ? ('按' + dim) : (grain ? ('按' + { year: '年', quarter: '季度', month: '月', week: '周', day: '日' }[grain]) : (period ? period.label : '分组'));
    return {
        period: period,
        grain: grain,
        dimColumn: dim,
        valueColumn: valueHit,
        aggFunc: _aggFunc(s),
        collapse: collapse,
        follow: follow,
        chartType: _chartType(s),
        label: label
    };
}

function _bucket(d, grain) {
    if (!d) return '';
    if (grain === 'year') return d.y + '年';
    if (grain === 'quarter') return d.y + '年Q' + Math.ceil(d.m / 3);
    if (grain === 'month') return d.y + '年' + d.m + '月';
    if (grain === 'day') return d.y + '-' + d.m + '-' + (d.day || 1);
    const dt = new Date(Date.UTC(d.y, d.m - 1, d.day || 1));
    const onejan = new Date(Date.UTC(d.y, 0, 1));
    const week = Math.max(1, Math.ceil((((dt - onejan) / 86400000) + onejan.getUTCDay() + 1) / 7));
    return d.y + '年第' + week + '周';
}

function _numericHeaders(headers, rows, skip) {
    const hs = headers || [];
    const rs = rows || [];
    return hs.filter(function (h) {
        if (!h || h === skip || /日期|时间|date|time/i.test(h)) return false;
        let seen = 0, ok = 0;
        for (let i = 0; i < Math.min(rs.length, 12); i++) {
            const v = _cellOf(rs[i], h, hs.indexOf(h));
            if (v == null || v === '') continue;
            seen++;
            if (_num(v) != null) ok++;
        }
        return seen > 0 && ok / seen >= 0.6;
    });
}

/** 先按季度留下行，再按日期粒度或某一列收成汇总表。收不成就保持原行。 */
function applyGroupView(headers, rows, plan) {
    const hs = Array.isArray(headers) ? headers.slice() : [];
    let rs = Array.isArray(rows) ? rows.slice() : [];
    if (!plan) return { headers: hs, rows: rs, note: '', changed: false };
    const notes = [];
    if (plan.period) {
        const filtered = filterRowsByQuarter(hs, rs, plan.period);
        if (filtered.applied) {
            rs = filtered.rows;
            notes.push('仅' + plan.period.label + '（' + plan.period.startMonth + '–' + plan.period.endMonth + '月，' + filtered.kept + '行）');
        } else if (filtered.empty) notes.push('没有' + plan.period.label + '，使用全部数据');
    }
    if (!plan.collapse) return { headers: hs, rows: rs, note: notes.join('，'), changed: notes.length > 0 };
    const dateCol = plan.grain ? findDateColumn(hs, rs) : '';
    const dim = plan.dimColumn || '';
    if (plan.grain && !dateCol && !dim) return { headers: hs, rows: rs, note: notes.join('，'), changed: notes.length > 0 };
    const keyName = plan.grain ? { year: '年份', quarter: '季度', month: '月份', week: '周', day: '日期' }[plan.grain] : dim;
    const valueCols = plan.valueColumn ? [plan.valueColumn] : _numericHeaders(hs, rs, dim || dateCol);
    const useCols = valueCols.length ? valueCols : [];
    const groups = new Map();
    rs.forEach(function (row) {
        let key = '';
        if (plan.grain && dateCol) {
            const d = parseCellDate(_cellOf(row, dateCol, hs.indexOf(dateCol)));
            if (!d) return;
            if (plan.period && plan.period.year && d.y !== plan.period.year) return;
            key = _bucket(d, plan.grain);
        } else if (dim) {
            const v = _cellOf(row, dim, hs.indexOf(dim));
            key = v == null || v === '' ? '' : String(v);
        }
        if (!key) return;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(row);
    });
    if (!groups.size) return { headers: hs, rows: rs, note: notes.join('，'), changed: notes.length > 0 };
    const outHeaders = [keyName].concat(useCols.length ? useCols : ['数量']);
    const outRows = [];
    groups.forEach(function (list, key) {
        const o = {};
        o[keyName] = key;
        if (useCols.length) {
            useCols.forEach(function (col) {
                o[col] = _reduce(list.map(function (row) { return _cellOf(row, col, hs.indexOf(col)); }), plan.aggFunc || 'sum');
            });
        } else o['数量'] = list.length;
        outRows.push(o);
    });
    notes.push('已按' + (plan.label || keyName) + '分成 ' + outRows.length + ' 组');
    return { headers: outHeaders, rows: outRows, note: notes.join('，'), changed: true, dateColumn: dateCol };
}

function attachGroupView(message, turn) {
    const sel = turn && turn.context && turn.context.selection;
    if (!sel || !Array.isArray(sel.headers) || !sel.headers.length || !Array.isArray(sel.rows)) return null;
    const plan = parseGroupPlan(message, sel.headers);
    if (!plan) return null;
    const view = applyGroupView(sel.headers, sel.rows, plan);
    if (plan.grain) plan.dateColumn = view.dateColumn || findDateColumn(sel.headers, sel.rows);
    const sheet = sel.sheetName || (turn.context.activeSheet && turn.context.activeSheet.sheet) || '';
    if (sheet) plan.sheet = sheet;
    turn.context.groupPlan = plan;
    turn.context.groupView = view;
    return plan;
}

function pptRowsFrom(turn, ctxSel) {
    const view = turn && turn.context && turn.context.groupView;
    if (view && view.changed) {
        return { headers: view.headers, rows: view.rows, note: view.note || '', sheetName: ctxSel && ctxSel.sheetName };
    }
    return {
        headers: ctxSel && ctxSel.headers,
        rows: ctxSel && ctxSel.rows,
        note: '',
        sheetName: ctxSel && ctxSel.sheetName
    };
}

/** 「根据这份数据生成一份 Q3 经营分析 PPT」→「Q3 经营分析」。抽不出就返回空串。 */
function pptReportTitle(message, period) {
    let t = String(message || '').trim();
    t = t.replace(/^(?:请|帮我|帮忙)?(?:根据|按照|按|用)(?:这份|当前|这个|选中的?)?(?:数据|表格|选区|工作表)/, '');
    t = t.replace(/^(?:请|帮我|帮忙)/, '');
    t = t.replace(/(?:帮我|请)?(?:生成|制作|做成|做|创建|出|来)(?:一份|一个|份)?/, '');
    t = t.replace(/(?:的)?(?:ppt|pptx|幻灯片|演示文稿|汇报(?:材料|文稿)?)$/i, '');
    t = t.replace(/^[，,：:\s]+|[，,。\s]+$/g, '').trim();
    if (t.length >= 2 && t.length <= 24) return t;
    if (period) return period.label + '经营分析';
    return '';
}

function shapePptArgs(message, headers, rows, sheetName, groupNote) {
    const period = parseQuarterPeriod(message);
    const title = pptReportTitle(message, period);
    const baseSource = sheetName ? (sheetName + ' 选区') : 'Excel 选区';
    let useHeaders = headers;
    let useRows = Array.isArray(rows) ? rows.slice() : [];
    let source = baseSource;
    let requirement = String(message || '');
    if (groupNote) {
        source += ' · ' + groupNote;
        requirement += '。趋势、对比。' + groupNote + '。';
    } else {
        const plan = parseGroupPlan(message, headers);
        const view = applyGroupView(headers, useRows, plan);
        if (view && (view.changed || view.note)) {
            useHeaders = view.headers;
            useRows = view.rows;
            if (view.note) {
                source += ' · ' + view.note;
                requirement += '。趋势、对比。' + view.note + '。';
            }
        } else if (period) {
            requirement += '。趋势、对比。' + period.label + '指公历' + period.startMonth + '–' + period.endMonth + '月。';
        }
    }
    return {
        reportTitle: title || ('数据分析报告 · ' + new Date().toLocaleDateString()),
        requirement: requirement,
        headers: useHeaders,
        rows: useRows,
        sourceLabel: source,
        period: period
    };
}

function _mkExecResult(o) {
    o = o || {};
    return {
        ok: o.ok === true,
        result: o.result,
        error: o.error || null,
        raw: (o.raw !== undefined) ? o.raw : null,
        steps: Array.isArray(o.steps) ? o.steps : undefined,
        // 步骤面板已完整覆盖主调用时置 true：orchestrator 跳过补裸工具卡（避免重复行）
        suppressToolCard: o.suppressToolCard === true,
        elapsedMs: (typeof o.elapsedMs === 'number') ? o.elapsedMs
            : (typeof o.durationMs === 'number' ? o.durationMs : 1)
    };
}

const McpDispatcher = {
    /**
     * 确定性匹配 + 直接执行（单工具，短路 Ollama）
     * 
     * @param {object} opts  { message, toolRegistry, sessionStore, turn, signal }
     * @returns {Promise<{hit: boolean, result?: object, score?: number, toolName?: string}>}
     */
        async tryDispatch(opts) {
        const { message, sessionStore, turn } = opts;
        if (!message || typeof message !== 'string') return { hit: false };

        // ── 第 0 步：纯关键词硬短路（绝不走 LLM / 匹配API / 注册表，100% 确定性）──
        // 只要用户消息命中已知的"计算机操作"口语词，就硬编码成对应 computer_ 工具并直接调后端。
        // 这是兜底中的兜底：即使 MatchCapability / toolRegistry / agent-loop 全部失效，
        // 截图这种最常用的操作也一定能跑通，不会再被 Ollama 幻觉截胡。
        const m_ = message.toLowerCase().trim();
        let hardTool = null;
        let hardArgs = {};
        let hardAfter = null; // 组合意图链：首工具成功后追加执行的工具
        // 【多任务规划·拆分后美化】"按X拆分成不同Sheet，然后分别对工作表美化"：
        // 拆分成功后在执行期按产物表（created/replaced）逐表链式美化。参数在执行期才能确定
        // （依赖拆分返回的新表名），故这里只挂标记，链骨架在首工具成功后现场搭。
        let splitThenFormat = false;
        let splitChainFmt = {};   // 链式美化参数：{} → 后端商务蓝预设；带 format/alignment → 用户点名的具体样式
        let splitChainLabel = '商务蓝美化预设';   // 注记用的样式名（主题名/指定样式/商务蓝预设）
        let skipRefine = false;   // 复合链场景参数已完全确定，跳过 LLM 语义补参
        // 【批量美化】"美化工作簿中的所有工作表"：枚举出的可见表清单。首表由主调用（hardTool）执行，
        // 其余表在执行期进链（batchBeautifyRest）。执行期才需要，故与 splitThenFormat 同 paradigm。
        let batchBeautifyRest = null;
        let batchBeautifyFmt = {};   // 批量美化样式参数：{} → 商务蓝预设；带 format/alignment → 用户点名的具体样式
        let batchBeautifyLabel = '商务蓝美化预设';   // 注记用的样式名
        let batchBeautifySkipped = [];   // 检测期就解析不出可靠范围的表（执行期汇总里点名，不静默丢）
        let batchSkipAll = false;   // 全部表都定位不到数据范围 → 跳过主调用，直接产出诚实的逐表点名汇报
                                    //（实测 2026-10-22 21:40：此场景旧版静默放行 LLM → 模型编造"已执行"被
                                    //  零工具守卫道歉，用户看到的是gaslight而非真相。空表/模块缺失都走这里）
        let batchEnumFail = false;  // excel_get_sheets 枚举失败/解析为空 → 诚实报错（22:04 事故：
                                    // 放行后 get_sheets 被通用兜底误选中，原始 JSON 倾倒进正文）
        let batchEnumDiag = '';     // 枚举失败的诊断头（raw/异常摘要）——写进用户可见报错，
                                    // 用户截图即可定位（用户会话的 hostLog 不落盘，实测 22:10）
        // 【主调用卡片抑制】批量美化/拆分后美化走"确定性步骤面板"，主调用本身已被步骤列表第一行
        // 覆盖（美化「销售二部」…其余表逐行）。再让 orchestrator 补一张裸工具卡就变成
        // `excel_format_range {} · 销售二部 · 1ms` 这种看不懂的重复行（实测 2026-10-22 22:31
        // 用户截图："没有任何改善"）。批量链场景置 true → orchestrator 跳过裸卡。
        let suppressMainToolCard = false;
        let markDupRows = false; // 「标记重复行」：先算出行号，再标黄，不删除
        // 宿主判定：Excel 专用的确定性硬短路（excel_format_range / apply_filter / check_errors /
        // clean_data）绝不能在 PPT/Word 宿主上触发，否则会把 word_/ppt_ 请求误派发到 excel_* 工具。
        // 宿主信息缺失时按 Excel 处理（保持既有行为）。
        const _isExcelHost = !(typeof window !== 'undefined' && window.SmartPilotHost && !window.SmartPilotHost.isExcel);
        // 入口留痕：宿主判定一旦出错，后面所有宿主专属分支都会走错（如把 ppt_ 请求当 Excel 处理）。
        hostLog('INFO', `tryDispatch 入口 host=${(typeof window !== 'undefined' && window.SmartPilotHost && window.SmartPilotHost.kind) || 'unknown'} isExcel=${_isExcelHost} msg=${String(message).slice(0, 60)}`);
        if (_isExcelHost && turn && turn.context && !turn.context.groupPlan) {
            try { attachGroupView(message, turn); } catch (eGrp) { ChatLog.warn?.('McpDispatcher', '分组计划跳过: ' + ((eGrp && eGrp.message) || eGrp)); }
        }
        const hitKw = list =>
            list.some(k => k && m_.includes(k.toLowerCase()));
        // 链式步骤的标题/详情由工具名数据驱动，避免为"打开→播放""切换→截图"各自硬编码显示文案。
        const stepMeta = (tool, args) => {
            const a = args?.app ? ('app=' + args.app) : JSON.stringify(args || {});
            switch (tool) {
                case 'computer_open_application': return { title: '打开应用', detail: tool + ' ' + a };
                case 'computer_focus_window': return { title: '切换窗口', detail: tool + ' ' + a };
                case 'computer_play_media': return { title: '聚焦窗口触发播放并验证', detail: tool + ' ' + a };
                case 'computer_capture_window': return { title: '截图该窗口', detail: tool + ' ' + a };
                case 'excel_aggregate': return { title: '按列分组', detail: tool + ' ' + a };
                case 'excel_date_group': return { title: '按时间分组', detail: tool + ' ' + a };
                case 'excel_create_chart': return { title: '用分组结果画图', detail: tool + ' ' + a };
                case 'excel_split_by_column': return { title: '按列拆分工作表', detail: tool + ' ' + a };
                case 'excel_get_used_range': return { title: '定位数据区域', detail: tool + ' ' + a };
                case 'excel_format_range': return { title: '美化表格', detail: tool + ' ' + a };
                default: return { title: tool, detail: tool + ' ' + a };
            }
        };
        // "切换到X窗口[，并截图]" / "截图X窗口"：置前目标窗口 或 按窗口区域截图（非整屏）。
        // 必须优先于下方通用"截图"分支，否则"截图"关键词会先命中 computer_screen_capture 整屏截走。
        const notWinWord = /工作簿|表格|文件|文档|对话|页面|视图|设置/;
        const fw = m_.match(/切换到([^\s，。！？、,.!?;；:：]{1,20}?)(?:窗口|应用|软件)/);
        const cw = m_.match(/(?:截图|截屏)(?:该|这个|此)?([^\s，。！？、,.!?;；:：]{1,20}?)(?:窗口|应用|软件)/);
        if (fw && String(fw[1] || '').trim() && !notWinWord.test(fw[1])) {
            const t = String(fw[1]).replace(/[，。！？、,.!?;；:：\s]+$/, '');
            hardTool = 'computer_focus_window'; hardArgs = { app: t };
            // "切到X并截图" → 链式：置前后按该窗口区域截图
            if (/截(?:图|屏)/.test(m_)) hardAfter = { tool: 'computer_capture_window', args: { app: t, returnFormat: 'path' } };
        } else if (cw && String(cw[1] || '').trim() && !notWinWord.test(cw[1])) {
            const t = String(cw[1]).replace(/[，。！？、,.!?;；:：\s]+$/, '');
            hardTool = 'computer_capture_window'; hardArgs = { app: t, returnFormat: 'path' };
        } else if (hitKw(['截屏','截图','屏幕截图','屏幕快照','截个屏','拍照','抓屏','录制屏幕','屏幕拍照','截一下屏','截图下来'])) {
            hardTool = 'computer_screen_capture';
            // 全屏截图 base64 可达几 MB，跨 WebView2 边界传输易被截断/无法渲染成 data URI。
            // 默认让后端把截图存成临时 PNG 文件并返回本地路径，前端用 file:// 引用，绕开超长 base64。
            if (!hardArgs.returnFormat) hardArgs.returnFormat = 'path';
        } else if (hitKw(['识别屏幕','识别屏幕文字','屏幕识别','截屏识别','ocr识别屏幕','识别截图里的字','识别图片文字'])) {
            hardTool = 'computer_screen_ocr';
        } else if (/^(?:退出|关闭|关掉|结束)(.{1,16}?)(?:应用|程序)?(?:吧|呀|吗)?$/.test(m_)) {
            // 退出/关闭应用：走 computer_close_application 真正杀进程，防止假确认。
            // 兜住网址/Excel 等特殊词另作处理，避免把"关闭工作簿"这类 Excel 语义抢走。
            const exM = m_.match(/^(?:退出|关闭|关掉|结束)(.{1,16}?)(?:应用|程序)?(?:吧|呀|吗)?$/);
            let exApp = exM ? String(exM[1]).replace(/[。！？、,.!?;；:：\s]+$/, '') : '';
            // 只对"真应用"放行（知名软件名或英文 exe/别名），绝不放行 Excel/文件/对话类语义，
            // 否则"关闭工作簿/关闭表格/退出对话"会被误当进程杀。
            const noKill = ['应用','程序','工作簿','工作表','表格','表格文件','文档','窗口','页面','话术','对话','助手','ai','文件','视图','模式'];
            const appLike = /^[a-z][a-z0-9.]*$/.test(exApp) &&
                !['excel','winword','powerpnt','excel.exe','notepad'].includes(exApp) ||
                ['音乐','qq音乐','浏览器','微信','wechat','chrome','edge','firefox','notepad','记事本','画图','mspaint','计算器','calc'].includes(exApp);
            if (exApp && !noKill.includes(exApp) && appLike) {
                hardTool = 'computer_close_application'; hardArgs = { app: exApp };
            }
        } else if (/打开([^\s，。！？、,.!?;；:：]{1,16}?)(?:并|然后|接着|再|，|,|、){0,2}(?:执行|开始|继续|点|去)?(?:播放|暂停)(?:音乐|歌曲|歌|一下)?/.test(m_)) {
            // 组合意图"打开X并播放音乐"：媒体键需应用已运行才有意义 → 先打开应用。
            // 连接词可重复 {0,2} 以兼容"，并执行播放"这类"逗号+连接词"写法。
            const cm = m_.match(/打开([^\s，。！？、,.!?;；:：]{1,16}?)(?:并|然后|接着|再|，|,|、){0,2}(?:执行|开始|继续|点|去)?(?:播放|暂停)(?:音乐|歌曲|歌|一下)?/);
            let comApp = cm ? String(cm[1]).replace(/[，。！？、,.!?;；:：]+$/, '') : '';
            const comIsFile = /\.(xlsx|xls|csv|docx?|pptx?|mp3|wav|flac|m4a|aac|txt|pdf|png|jpe?g|bmp|gif)$/i.test(comApp);
            if (comApp && !comIsFile) { hardTool = 'computer_open_application'; hardArgs = { app: comApp }; }
            if (comApp && !comIsFile && /播放/.test(m_)) {
                // 打开成功后用 computer_play_media：聚焦窗口→OCR判已在播→空格切换→OCR复验"正在播放"。
                // 取代"坐标点播放按钮"——OCR坐标点QQ音乐的播放控件不可靠（可能点到单曲/非播放钮）。
                hardAfter = { tool: 'computer_play_media', args: { app: comApp } };
            }
        } else if (hitKw(['播放音乐','暂停音乐','继续播放','接着播放','播放mp3','播放暂停','播放一下'])) {
            // 媒体播放/暂停（系统级媒体键，对 QQ音乐/网易云/浏览器媒体生效）
            hardTool = 'computer_key_press'; hardArgs = { key: 'media_play_pause' };
        } else if (hitKw(['下一首','切歌','下一曲','下首歌','下一首歌'])) {
            hardTool = 'computer_key_press'; hardArgs = { key: 'media_next' };
        } else if (hitKw(['上一首','上一曲','上首歌','上一首歌'])) {
            hardTool = 'computer_key_press'; hardArgs = { key: 'media_prev' };
        } else if (hitKw(['打开浏览器','启动浏览器','打开chrome','打开edge','打开浏览'])) {
            hardTool = 'computer_open_application'; hardArgs = { app: 'chrome' };
        } else if (hitKw(['打开excel','打开表格','启动excel','打开工作簿'])) {
            hardTool = 'computer_open_application'; hardArgs = { app: 'excel' };
        } else {
            // "点击/双击桌面上的 X 图标" → 语义等价于启动该应用（打开桌面快捷方式），非真实像素级点击
            const dm = m_.match(/(?:点击|双击|点开|启动)(?:桌面上的|桌面的)([^\s，。！？、,.!?;；:：]{1,16})(?:图标|快捷方式)/);
            if (dm && !/(chrome|浏览器|edge|工作簿|文件|文件夹|表格|设置)/i.test(dm[1])) {
                hardTool = 'computer_open_application'; hardArgs = { app: dm[1] };
            } else {
                // 通用"打开任意应用"：非 浏览器/edge/excel/文件类 的"打开X" → 交给后端按 app 名解析（
                // 开始菜单快捷方式 / 已知别名 / App Paths）。不劫持文件/工作簿类打开，避免跟 Excel 抢。
                const om = m_.match(/打开([^\s，。！？、,.!?;；:：]{1,24})/);
                let app = om ? String(om[1]).replace(/[，。！？、,.!?;；:：]+$/, '') : '';
                const isFileOpen = /\.(xlsx|xls|csv|docx?|pptx?|txt|pdf|png|jpe?g|bmp|gif|zip|rar|7z)$/i.test(app);
                const isKnownVerb = /(chrome|浏览器|edge|工作簿|文件|文件夹|表格|word|ppt|工作表|sheet|资源管理器|应用商店|设置)/i.test(app);
                if (app && !isFileOpen && !isKnownVerb) {
                    hardTool = 'computer_open_application'; hardArgs = { app };
                }
            }
        }

        // ── Excel 宿主：导出/另存为 CSV → 直调 excel_export_csv（必须放在下面"交规划"之前）──
        // 背景（实测 2026-10-28 19:04）：用户说"将当前工作表导出为 CSV 并保存到桌面"，
        // Excel 侧没有对应硬短路 → 交给 agent-loop，弱模型第 1 轮先调 excel_save_workbook
        // 存了个 xlsx，第 2 轮（总结轮，tools 已按 toolsExecuted 锁死不下发）只能写下
        // "调用 excel_export_csv 工具导出 CSV 文件"的永动机计划 —— CSV 永远没导出。
        // 导出 CSV 是纯文件产出类动作（与 Word/PPT 导出 PDF 同性质）：不需要读数据、
        // 不需要 LLM 参与；不猜 sheet（后端默认活动表）、filePath 留空由后端落盘桌面。
        // 需要先做别的活（筛选/汇总/美化/图表…）或批量多表时仍交规划，避免只导出就没下文。
        if (!hardTool && _isExcelHost) {
            const EXPORT_VERB = RC().EXPORT_VERB;
            const CSV_KW = /(csv|逗号分隔|文本文件)/i;
            // 注意不能把裸"格式"当排除词——"导出为 CSV 格式"是最常见说法；
            // 只有"格式化/格式设置/条件格式…"这类真要做格式的才交规划。
            const CSV_NEEDS_PLAN = /(筛选|汇总|统计|分析|图表|画图|透视|美化|格式化|格式设置|条件格式|数字格式|日期格式|清洗|合并|拆分|排序|去重|删除|每个|每张|所有|全部|各个|分别|批量)/;
            if (EXPORT_VERB.test(message) && CSV_KW.test(message) && !CSV_NEEDS_PLAN.test(message)) {
                const sel = (turn && turn.context && turn.context.selection) || null;
                const _sheet = (sel && sel.sheetName)
                    || (turn && turn.context && turn.context.activeSheet && turn.context.activeSheet.sheet) || '';
                hardTool = 'excel_export_csv';
                hardArgs = {};
                if (_sheet) hardArgs.sheet = String(_sheet);
                ChatLog.info?.('McpDispatcher', `🎯 [excel] 导出 CSV 硬短路 → ${hardTool} sheet=${hardArgs.sheet || '(活动表)'}`);
                hostLog('INFO', `🎯 [excel] 导出 CSV 硬短路 → ${hardTool} sheet=${hardArgs.sheet || '(活动表)'}`);
            }
        }

        // Excel 上的任务不再关键词硬短路。合并、美化、拆分、筛选都会在这里把整句做完，
        // 后半句（「并设置成中国红」）进不了规划。计算机操作（截图、打开应用）仍走上面的硬路径。
        // allowHardDispatch 只给测试覆盖旧引擎，线上调用不会带这个参数。
        if (_isExcelHost && !hardTool && opts.allowHardDispatch !== true) {
            let styleTarget = '';
            try {
                styleTarget = (App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.semanticStyleTarget === 'function')
                    ? (App.ChatV2.turnGoal.semanticStyleTarget(message) || '') : '';
            } catch (_) { styleTarget = ''; }
            ChatLog.info?.('McpDispatcher', '取消硬短路：Excel 任务交给模型规划' + (styleTarget ? '（外观「' + styleTarget + '」）' : ''));
            hostLog('INFO', '取消硬短路：Excel 任务交给模型规划' + (styleTarget ? ' style=' + styleTarget : ''));
            return { hit: false, deferToPlanner: true, reason: 'llm-plan', styleTarget: styleTarget };
        }

        // ── 【PPT 生成意图硬短路】────────────────────────────────────────
        // "做成 PPT / 生成 PPT / 把数据做成 PPT" 这类明确意图，100% 确定性直调 ppt_generate，
        // 绕开 LLM / agent-loop 多轮规划及其"重复调用同一工具"风险（实测：小模型每轮复读
        // ppt_generate，连生成多份文件、UI 连弹多张相同卡片、反复复述"已完成"）。
        // 仅对"简单生成"生效；含"逐页 / 每页 / 自定义结构"等自定义诉求时放行 agent-loop
        // （走 ppt_project_start → ppt_add_*_slide → ppt_project_build 分步链路）。
        // 仅当有可用选区数据时才硬短路——无选区的纯结构 PPT 交给 agent-loop（让其规划或反问），
        // 避免硬短路空参裸调后端失败。命中后复用下方统一执行块（含 ppt_ 选区注入与卡片渲染）。
        if (!hardTool) {
            const pptGenKw = /(做|生成|制作|创建|出|转成|转换成|弄|搞|整|来|写|给(我|份|个)?).{0,14}(ppt|幻灯片|演示文稿|汇报(材料|文稿)?|pptx|简报|keynote)/i;
            // 看板/大屏类（如"生成 ppt 大看板"）不进硬短路单工具，放行 agent-loop 走
            // ppt_project_start → ppt_add_*_slide → ppt_project_build 分步链，从而渲染真实多步卡片。
            const pptFreeBuild = /逐页|每一页|每页|自定义结构|分页|封面.*目录|指定.*页面|分章节|目录页.*正文页|自己设计.*页|分步构建|看板|仪表盘|驾驶舱|大屏|大看板|数据大屏|数据面板|dashboard|databoard/i;
            if (pptGenKw.test(message) && !pptFreeBuild.test(m_)) {
                const ctxSel = turn && turn.context ? turn.context.selection : null;
                const hasSel = ctxSel && ctxSel.success !== false &&
                    Array.isArray(ctxSel.headers) && ctxSel.headers.length > 0 &&
                    Array.isArray(ctxSel.rows) && ctxSel.rows.length > 0;
                if (hasSel) {
                    const pptSel = pptRowsFrom(turn, ctxSel);
                    const shaped = shapePptArgs(message, pptSel.headers, pptSel.rows, pptSel.sheetName, pptSel.note);
                    hardTool = 'ppt_generate';
                    hardArgs = {
                        reportTitle: shaped.reportTitle,
                        requirement: shaped.requirement,
                        autoDesign: true,
                        sourceLabel: shaped.sourceLabel,
                        headers: shaped.headers,
                        rows: shaped.rows
                    };
                    ChatLog.info?.('McpDispatcher', `🎯 PPT 硬短路命中 ppt_generate（选区 ${ctxSel.headers.length}列 × ${shaped.rows.length}行，标题 ${shaped.reportTitle}）`);
                } else {
                    ChatLog.info?.('McpDispatcher', '[PPT 硬短路] 无可用选区数据，放行通用管道');
                }
            }
        }

        // ── 非 Excel 宿主（PPT/Word）：选中表格的格式化请求 → 直调该宿主的"选中表格格式化"工具 ──
        // 用户选中表格后说"设置对齐/颜色/字体/美化一下"，不该再反问行列号。
        // 省略 row/col → 后端整表生效；本处只解析 align/表头色等，未提及项留空由后端处理。
        if (!hardTool && !_isExcelHost) {
            const FMT_VERBS_HOST = RC().FMT_VERBS_HOST;
            if (FMT_VERBS_HOST.some(k => m_.includes(k))) {
                const _ctx = (turn && turn.context) || {};
                const _selObj = _ctx.selectedObject
                    || (_ctx.selection && (_ctx.selection.objectType === 'table' || _ctx.selection.selectionType === 'table') ? _ctx.selection : null);
                // 判定"选中表格"：显式 selectionType，或带回行列数，或摘要含"表格/单元格"。
                // 放宽是为了兼容不同宿主/不同版本桥返回字段的差异——否则一旦字段名不一致就会静默落空，
                // 退化成"模型反过来问用户要行号/列号"（正是截图里的症状）。
                const _isTableSel = !!_selObj && (
                    _selObj.selectionType === 'table' ||
                    (Number(_selObj.rowCount) > 0 && Number(_selObj.colCount) > 0) ||
                    /表格|单元格/.test(String(_selObj.summary || ''))
                );
                if (_isTableSel) {
                    const _fam = (window.SmartPilotHost && window.SmartPilotHost.isPowerPoint) ? 'ppt' : 'word';
                    hardTool = _fam + '_format_selected_table_cell';
                    hardArgs = buildHostTableFormatArgs(message);
                    ChatLog.info?.('McpDispatcher', `🎯 [${_fam}] 选中表格格式化硬短路 → ${hardTool} args=${JSON.stringify(hardArgs)}`);
                    hostLog('INFO', `🎯 [${_fam}] 选中表格格式化硬短路 → ${hardTool} args=${JSON.stringify(hardArgs)}`);
                } else {
                    // 留痕：下次若仍"不执行"，可直接据此判断是识别失败还是分支未生效
                    const _diag = 'selectedObject=' + (_ctx.selectedObject
                        ? ('{type=' + (_ctx.selectedObject.selectionType || '') + ',rows=' + (_ctx.selectedObject.rowCount || '') + ',cols=' + (_ctx.selectedObject.colCount || '') + ',summary=' + String(_ctx.selectedObject.summary || '').slice(0, 40) + '}')
                        : 'none')
                        + '，selection=' + (_ctx.selection ? ('{objectType=' + (_ctx.selection.objectType || '') + ',selectionType=' + (_ctx.selection.selectionType || '') + '}') : 'none')
                        + '，ctxKeys=' + Object.keys(_ctx).slice(0, 12).join('/');
                    ChatLog.info?.('McpDispatcher', `[非 Excel 格式化] 未识别到已选表格，放行管道（${_diag}）`);
                    hostLog('WARN', `[非 Excel 格式化] 未识别到已选表格，放行管道（${_diag}）`);
                }
            } else {
                hostLog('INFO', '[非 Excel 宿主] 消息未命中格式类动词，不进入选中表格硬短路');
            }
        }

        // ── 非 Excel 宿主（PPT/Word）：文档级"导出/另存为 PDF（图片）"→ 直调该宿主导出工具 ──
        // 背景（实测 2026-10-06 00:48）：Word 里说"将当前文档导出为 PDF"，此处无规则 →
        // 落到 agent-loop 后小模型认为"文档中没有任何可读取的内容"直接反问用户要内容，
        // 导出工具一次都没调（用户看到的就是"导出PDF要什么可读取的文字内容"）。
        // 导出是 100% 确定性的文件产出动作：不需要读文档、不需要 LLM 参与。
        // 命中即直调，filePath 留空交后端落盘桌面（app-mcp 的 _fillOptionalExportPathIfNeeded
        // 会先把真实桌面路径补进参数，工具卡片"参数"区可见完整路径）。
        if (!hardTool && !_isExcelHost) {
            const _isPptHost = !!(typeof window !== 'undefined' && window.SmartPilotHost && window.SmartPilotHost.isPowerPoint);
            const _fam = _isPptHost ? 'ppt' : 'word';
            const EXPORT_VERB = RC().EXPORT_VERB;
            const PDF_KW = RC().PDF_KW;
            // 更具体的目标（PDF > 图片）优先；未提及目标时不猜（放行下游，避免把"导出数据"误当导出文件）
            if (EXPORT_VERB.test(message) && PDF_KW.test(message)) {
                hardTool = _fam + '_export_pdf';
                hardArgs = {};
                ChatLog.info?.('McpDispatcher', `🎯 [${_fam}] 导出 PDF 硬短路 → ${hardTool}（filePath 留空，后端落盘桌面）`);
                hostLog('INFO', `🎯 [${_fam}] 导出 PDF 硬短路 → ${hardTool}`);
            } else if (!hardTool && _isPptHost && EXPORT_VERB.test(message) && /(图片|截图|image|png|jpe?g)/i.test(message)) {
                hardTool = 'ppt_export_image';
                hardArgs = {};
                ChatLog.info?.('McpDispatcher', `🎯 [ppt] 导出图片硬短路 → ${hardTool}`);
                hostLog('INFO', `🎯 [ppt] 导出图片硬短路 → ${hardTool}`);
            }
        }

        // ── 单元格格式化动词硬短路（加粗/居中/倾斜/边框/字号/字体/填充色/行高/列宽/自动换行…）──
        // 与 computer_/ppt 第 0 步同理：100% 确定性，绕开 LLM。skill 关闭时仍走此确定性直调。
        // 目标区域优先级：① 消息显式区域（A1:D1）> ② 当前选区（"表头/首行"取首行，否则整块）
        //   > ③ 纯格式化命令且无区域 → 留空 range 交后端"实时选区"兜底（后端已去掉 range="A1"
        //     魔法、UsedRange 兜底需显式 autoRange，空 range 安全，不会误格式化整表）。
        // 含分析/复合操作意图的消息不走短路，放行下游多步骤/LLM。
        // 【仅 Excel】PPT/Word 由上面的分支处理，避免 excel_format_range 误派发到非 Excel 宿主。
        // 【批量排他】"给所有工作表加边框"这类多目标消息不进本块（本块只做单表/选区），
        // 让下面的批量美化硬短路枚举可见表逐表处理——否则会退化成"只格式化当前选区"。
        if (!hardTool && _isExcelHost && !isBatchBeautifyMessage(message) && isBareTableBeautify(message)) {
            const sel = (turn && turn.context && turn.context.selection) || null;
            const selRange = sel && (sel.address || sel.range);
            const selIsRange = !!selRange && /[:：]/.test(String(selRange).replace(/^[^!]*!/, ''));
            const sheet = (sel && sel.sheetName) || (turn && turn.context && turn.context.activeSheet && turn.context.activeSheet.sheet) || '';
            hardTool = 'excel_format_range';
            hardArgs = { format: '{}', beautifyPreset: true };
            if (selIsRange) hardArgs.range = String(selRange).replace(/^[^!]*!/, '');
            else hardArgs.autoRange = true;
            if (sheet) hardArgs.sheet = sheet;
            ChatLog.info?.('McpDispatcher', `🎯 美化表格走预设 excel_format_range range=${hardArgs.range || '(UsedRange)'}`);
        }

        if (!hardTool && _isExcelHost && !isBatchBeautifyMessage(message)) {
            const FMT_VERBS = RC().FMT_VERBS_EXCEL;
            if (FMT_VERBS.some(k => m_.includes(k))) {
                const rng = extractExcelRange(message);
                const fmtArgs = buildFormatArgs(message);
                // skill 关闭时仍要能确定性执行：无显式区域但有当前选区 → 用选区兜底直调。
                // "表头/标题行/首行"等词 → 命中选区首行；其余 → 整块选区。
                // 无显式区域也无前端选区时，仅"纯格式化命令"才兜底为"后端实时选区"直调：
                // 后端已移除 range="A1" 魔法、UsedRange 兜底需显式 autoRange → 空 range 安全，
                // 不会误格式化成整表。含分析/复合操作意图的消息不在此短路，放行多步骤/LLM。
                const _PURE_FMT = /(加粗|加斜|倾斜|斜体|居中|边框|字体颜色|字体|字号|填充|背景色|底色|底纹|高亮|行高|列宽|自动换行|清除格式|设置格式|格式化|单元格格式|对齐|表头|首行|标题行)/;
                const _NON_PURE_FMT = /(汇总|聚合|统计|分析|图表|画图|绘制|透视|对比|环比|同比|占比|趋势|分组|分类汇总|明细|洞察|复制|写回|写入|新建|粘贴|导出|移动|追加|拆分|合并|转换|生成|插入|删除)/;
                let effRange = rng;
                let deferBackendSel = false;   // 无区域：交后端"实时选区"兜底
                let wantAutoRange = false;     // 无区域但用户要"整张表"：交后端自动推断 UsedRange
                if (!effRange && fmtArgs) {
                    const sel = (turn && turn.context && turn.context.selection) || null;
                    const selRange = sel && (sel.address || sel.range);
                    if (selRange) {
                        const HEADER_WORDS = /表头|标题行|首行|第一行|表头行|header/i;
                        effRange = HEADER_WORDS.test(m_)
                            ? (firstRowRangeOf(selRange) || selRange)
                            : selRange;
                    } else if (_PURE_FMT.test(m_) && !_NON_PURE_FMT.test(m_)) {
                        if (/整张表|整个表|全表|整表|整个工作表|整张工作表|所有数据|全部数据/.test(m_)) {
                            wantAutoRange = true;
                        } else {
                            deferBackendSel = true;
                        }
                    }
                }
                if (fmtArgs && (effRange || deferBackendSel || wantAutoRange)) {
                    // 【多要求排他】确定性参数覆盖不到全部格式要求时，放弃硬短路 → 放行 agent-loop 多步处理。
                    // 否则未覆盖的要求（列宽 / 解析不出的底色 / 笼统的"对齐方式"…）会被静默丢弃，
                    // 而 UI 仍显示"任务完成"，用户投诉的正是这种"只做了 1 条却说做完了"。
                    const _missAspects = uncoveredFormatAspects(message, fmtArgs);
                    if (_missAspects.length > 0) {
                        ChatLog.info?.('McpDispatcher', `🎯 格式化硬短路放弃：确定性参数未覆盖 [${_missAspects.join('、')}] → 放行 agent-loop 多步处理`);
                        hostLog('INFO', `格式化硬短路放弃（多要求排他）：未覆盖 [${_missAspects.join('、')}] → 放行下游`);
                        return { hit: false, formatMultiGuard: true };
                    }
                    hardTool = 'excel_format_range';
                    hardArgs = Object.assign({}, fmtArgs);
                    if (effRange) hardArgs.range = effRange;
                    if (wantAutoRange) hardArgs.autoRange = true;
                    // 后端 GetArg(Of String) 读取 format/alignment，必须传 JSON 字符串而非对象
                    if (hardArgs.format && typeof hardArgs.format === 'object') hardArgs.format = JSON.stringify(hardArgs.format);
                    if (hardArgs.alignment && typeof hardArgs.alignment === 'object') hardArgs.alignment = JSON.stringify(hardArgs.alignment);
                    ChatLog.info?.('McpDispatcher', `🎯 格式化动词硬短路命中 excel_format_range range=${effRange || '(后端实时选区)'} args=${JSON.stringify(hardArgs)}`);
                    hostLog('INFO', `🎯 格式化动词硬短路命中 excel_format_range range=${effRange || '(后端实时选区)'} args=${JSON.stringify(hardArgs)}`);
                }
            }
        }

        // ── 筛选/过滤条件硬短路（"筛选 订单金额>3000 且 区域=华东"）──
        // 同 computer_/ppt/格式化：100% 确定性直调 excel_apply_filter，绕开 LLM。
        // 实测把"筛选"交给 LLM 会被演绎成"读全表 + 写分析报告"（15s / 16.7k token），
        // 根本不执行真正的筛选；故此处解析条件后确定性直调。
        // 条件键保持中文列名（后端按表头解析成列，支持"区域"→"销售区域"这类包含匹配）；
        // 区域优先级：① 消息显式区域 > ② 当前选区 > ③ 整张已用区域（autoRange）。
        // 含聚合/分析/排序/图表等复合意图时放行，交由多步骤/Python 管道处理。
        if (!hardTool && _isExcelHost) {
            const FILTER_KW = RC().FILTER_KW;
            if (FILTER_KW.test(m_)) {
                const NON_FILTER = /(汇总|聚合|统计|分析|图表|画图|绘制|透视|对比|环比|同比|占比|趋势|分组|分类汇总|明细|洞察|求和|平均|计数|最大|最小|排序|升序|降序|去重|删除|清空|清除筛选|取消筛选|撤销筛选)/;
                const criteria = parseFilterCriteria(message);
                if (Object.keys(criteria).length > 0 && !NON_FILTER.test(m_)) {
                    const rng = extractExcelRange(message);
                    const sel = (turn && turn.context && turn.context.selection) || null;
                    const selRange = sel && (sel.address || sel.range);
                    // 选区若是"单格"（如 A1）说明用户只是光标停在那儿，并非要筛这一个格 → 直接用整表；
                    // 只有多格选区（"A1:G201"）才算用户明确框选的目标区域。
                    const _selIsRange = !!selRange && /[:：]/.test(String(selRange).replace(/^[^!]*!/, ''));
                    hardTool = 'excel_apply_filter';
                    hardArgs = { criteria: JSON.stringify(criteria) };
                    if (rng) hardArgs.range = rng;
                    else if (_selIsRange) hardArgs.range = selRange;
                    else hardArgs.autoRange = true;   // 无有效区域：后端用 UsedRange 推断整表
                    ChatLog.info?.('McpDispatcher', `🎯 筛选硬短路命中 excel_apply_filter range=${hardArgs.range || '(后端 UsedRange)'} criteria=${JSON.stringify(criteria)}`);
                } else if (Object.keys(criteria).length === 0) {
                    ChatLog.info?.('McpDispatcher', '[筛选硬短路] 未解析出条件，放行通用管道');
                }
            }
        }

        // ── 错误检查硬短路（"检查/检测是否包含错误"）──
        // 100% 确定性直调 excel_check_errors（后端 COM SpecialCells 一次性定位 #DIV/0!、#N/A 等
        // 公式错误单元格），绕开 LLM。实测交给 agent-loop：LLM 每次都选 excel_read_range 自己
        // "看数据找异常"，#DIV/0! 被读成数字后报"未发现错误/出现异常最大值"，完全不可用。
        // 区域优先级：① 消息显式区域 > ② 当前多格选区 > ③ 整表 UsedRange（不传 range，后端默认）。
        // 含统计/图表/清洗等复合意图时不短路，放行通用管道。
        if (!hardTool && _isExcelHost) {
            const CHECK_KW = RC().CHECK_KW;
            const ERR_KW = /(错误|出错|error|#DIV|#N\/A|#VALUE|#REF|公式问题)/i;
            const NON_CHECK = /(汇总|聚合|统计|求和|平均|计数|图表|画图|绘制|透视|排序|筛选|去重|清洗|美化|格式|写入|填写|删除|复制|粘贴|导出|合并|拆分|趋势|预测|环比|同比|占比|分组)/;
            if (CHECK_KW.test(m_) && ERR_KW.test(m_) && !NON_CHECK.test(m_)) {
                const rng = extractExcelRange(message);
                const sel = (turn && turn.context && turn.context.selection) || null;
                const selRange = sel && (sel.address || sel.range);
                // 单格选区视为"光标停在那"，不算明确目标 → 用整表；多格选区才算用户框选的检查范围
                const _selIsRange = !!selRange && /[:：]/.test(String(selRange).replace(/^[^!]*!/, ''));
                hardTool = 'excel_check_errors';
                hardArgs = {};
                if (rng) hardArgs.range = rng;
                else if (_selIsRange) hardArgs.range = selRange;
                const _sheet = (sel && sel.sheetName) || (turn && turn.context && turn.context.activeSheet && turn.context.activeSheet.sheet) || '';
                if (_sheet) hardArgs.sheet = _sheet;
                ChatLog.info?.('McpDispatcher', `🎯 错误检查硬短路命中 excel_check_errors range=${hardArgs.range || '(整表 UsedRange)'} sheet=${_sheet || '(活动表)'}`);
            }
        }

        // ── 清洗硬短路（"清洗这份数据/去重/删空行/填充空值/规范日期"）──
        // 100% 确定性直调 excel_clean_data / excel_remove_duplicates，绕开 LLM。
        // 实测交给 agent-loop："清洗这份数据"被演绎成"读数据 + 写数据质量报告"（40s / 6.7k token），
        // 根本不执行清洗；且 read_range 读到含合并标题行的表会产出错误分析。
        // 操作映射：明确只要预览、或「不删除 / 只高亮」→ excel_remove_duplicates（只算不写）；
        // 其余（含「移除重复行」「去重」「清洗这张表」）→ excel_clean_data 就地写回。
        // excel_remove_duplicates 按设计只返回唯一行，不改工作表。用户说「移除」却被送到它，
        // 卡片就写成「仅计算，未写回 Excel」（实测「明确了移除却仅计算」）。
        // excel_clean_data 默认 dryRun=true，不传 writeBack 时工作表也不动。
        if (!hardTool && _isExcelHost) {
            const CLEAN_KW = /(清洗|清理|去重|删除重复|移除重复|重复行|重复数据|空行|空白行|去空格|多余空格|填充空|空值|补全空|规范日期|日期(格式)?(统一|规范)|规范数字|数字(格式)?(统一|规范))/;
            const NON_CLEAN = /(图表|画图|绘制|透视|看板|大屏|汇总|聚合|统计|求和|平均|计数|排序|筛选|导出|发送|邮件|ppt|分析报告)/i;
            if (CLEAN_KW.test(m_) && !NON_CLEAN.test(m_)) {
                const rng = extractExcelRange(message);
                const sel = (turn && turn.context && turn.context.selection) || null;
                const selRange = sel && (sel.address || sel.range);
                // 单格选区视为光标停留，不算明确目标 → 交后端 UsedRange 兜底
                const _selIsRange = !!selRange && /[:：]/.test(String(selRange).replace(/^[^!]*!/, ''));
                const _sheet = (sel && sel.sheetName) || (turn && turn.context && turn.context.activeSheet && turn.context.activeSheet.sheet) || '';
                const wantsDedupe = /(去重|删除重复|移除重复|重复行|重复数据)/.test(m_);
                const previewOnly = /(只预览|仅预览|先别改|不要修改|别改表|先看看|看看效果)/.test(m_);
                // 「不删除」里的「删除」不能算成要删行。
                const _markText = m_.replace(/不删除|不要删除|别删|勿删|不删/g, '');
                const explicitDelete = /(删除|移除|去掉|清除|删掉)/.test(_markText);
                const explicitMark = /(仅高亮|只高亮|只标记|仅标记|高亮标记|标黄|标记重复|标出重复|标注重复|高亮重复)/.test(m_)
                    || (/标记|高亮|标出|标注/.test(m_) && /重复/.test(m_));
                const markOnly = (explicitMark && !explicitDelete) || /(不删除|勿删|别删|不要删除)/.test(m_);
                const otherClean = /(空行|空格|填充|规范)/.test(m_);
                // 「标记重复行」是标黄，不是删行。excel_remove_duplicates 只算行号，算完再 format 标黄。
                // 只问「别改表 / 先看看」才停在纯计算。说了移除、去重、清洗就要删行写回。
                if (wantsDedupe && !otherClean && markOnly && !previewOnly) {
                    hardTool = 'excel_remove_duplicates';
                    hardArgs = { markOnly: true, hasHeaderRow: 'true' };
                    markDupRows = true;
                } else if (wantsDedupe && !otherClean && (previewOnly || markOnly)) {
                    hardTool = 'excel_remove_duplicates';
                    hardArgs = {};
                } else {
                    hardTool = 'excel_clean_data';
                    hardArgs = {};
                    const ops = [];
                    if (/(空行|空白行|删(除)?空)/.test(m_)) ops.push('removeEmptyRows');
                    if (/空格/.test(m_)) ops.push('trimSpaces');
                    if (/(填充|空值|补全空)/.test(m_)) ops.push('fillEmpty');
                    if (/(规范日期|日期(格式)?(统一|规范))/.test(m_)) ops.push('normalizeDates');
                    if (/(规范数字|数字(格式)?(统一|规范))/.test(m_)) ops.push('normalizeNumbers');
                    if (/(全半角|大小写|文本规范)/.test(m_)) ops.push('standardizeText');
                    if (wantsDedupe) ops.push('removeDuplicates');
                    // 「清洗这张表」没点名操作时，整行重复也要查出来。
                    // 旧默认只有去空格+删空行，第 16、17 行这种完全相同的记录会被放过。
                    if (ops.length === 0) ops.push('trimSpaces', 'removeDuplicates', 'removeEmptyRows');
                    hardArgs.operations = JSON.stringify(ops);
                    if (/填充.*[^01]0|按零填充|填零/.test(m_)) hardArgs.fillMethod = 'zero';
                    if (!previewOnly && !markOnly) {
                        hardArgs.dryRun = false;
                        hardArgs.writeBack = true;
                    }
                }
                if (rng) hardArgs.range = rng;
                else if (_selIsRange) hardArgs.range = selRange;
                if (_sheet) hardArgs.sheet = _sheet;
                ChatLog.info?.('McpDispatcher', `🎯 清洗硬短路命中 ${hardTool} range=${hardArgs.range || '(后端 UsedRange)'} ops=${hardArgs.operations || '(默认基础清洗)'}`);
            }
        }

        // ── 「移除销售一部数据」：删掉等于该值的整行，不能只读表后声称已删除 ──
        if (!hardTool && _isExcelHost) {
            const removeValue = extractRemoveRowValue(message);
            if (removeValue) {
                const sel = (turn && turn.context && turn.context.selection) || null;
                const _sheet = (sel && sel.sheetName) || (turn && turn.context && turn.context.activeSheet && turn.context.activeSheet.sheet) || '';
                hardTool = 'excel_delete_row';
                hardArgs = { matchValue: removeValue, deleteMatchingRows: true };
                if (_sheet) hardArgs.sheet = _sheet;
                ChatLog.info?.('McpDispatcher', `🎯 按值删行硬短路 excel_delete_row value=${removeValue}`);
            }
        }

        // ── 「标记王五 / 把李四标黄」──
        // 不是缺工具：excel_apply_conditional_format 的 highlight + equal 就能把等于该值的单元格涂黄。
        // 「标记王五」不含「重复」，清洗硬短路不接；意图分析把它判成 free_chat 再追问（实测未调用工具）。
        // 【批量排他】"把所有工作表改成蓝色"这类多目标涂色不进本块（本块只涂当前选区），
        // 让批量美化块枚举可见表逐表处理；extractMarkLiteral 内部也已排除主题词。
        if (!hardTool && _isExcelHost && !isBatchBeautifyMessage(message)) {
            const markLiteral = extractMarkLiteral(message);
            if (markLiteral) {
                const sel = (turn && turn.context && turn.context.selection) || null;
                const selRange = sel && (sel.address || sel.range);
                const selIsRange = !!selRange && /[:：]/.test(String(selRange).replace(/^[^!]*!/, ''));
                const _sheet = (sel && sel.sheetName) || (turn && turn.context && turn.context.activeSheet && turn.context.activeSheet.sheet) || '';
                let range = '';
                if (selIsRange) range = String(selRange).replace(/^[^!]*!/, '');
                if (!range) {
                    try {
                        const got = await callLocalToolWithFallback('excel_get_used_range', JSON.stringify(_sheet ? { sheet: _sheet } : {}));
                        range = _usedRangeAddr(got && got.raw) || '';
                    } catch (eUr) { /* 拿不到已用范围时交给后端选区 → UsedRange */ }
                }
                const fill = markFillColor(message);
                hardTool = 'excel_apply_conditional_format';
                hardArgs = {
                    formatType: 'highlight',
                    options: JSON.stringify({ operator: 'equal', value: markLiteral, color: fill.hex }),
                    markLiteral: markLiteral,
                    markColorName: fill.name
                };
                if (range) hardArgs.range = range;
                if (_sheet) hardArgs.sheet = _sheet;
                ChatLog.info?.('McpDispatcher', `🎯 标记字面值硬短路 excel_apply_conditional_format value=${markLiteral} range=${range || '(后端回退)'}`);
            }
        }

        // ── 合并工作表硬短路（"将选中工作表进行数据合并 / 合并这几张表 / 合到一张表"）──
        // 100% 确定性直调 excel_merge_sheets，绕开 LLM / agent-loop。
        // 实测（2026-10-07 19:26，用户连续 4 次）："将选中工作表进行数据合并" 被自由发挥成
        // "读大客户部 → 聚合 → 画饼图"，用户要的合并**一次都没跑**。
        // 根因有两层：① "合并"既是 excel_merge_cells（合并单元格）的动词、也是数据合并的动词，
        // 工具瘦身（filterToolsForMessage）按"合并"只保留了 merge_cells 族 → excel_merge_sheets
        // 被剔除，模型手上根本没有合并工具；② dispatcher 只有清洗/筛选/检查/格式化硬短路，
        // 没有"合并工作表"，于是整条请求交给 LLM 自由发挥。
        // 判定 = 合并动词 + 工作表级对象，且排除单元格合并 / 多文件合并 / 复合意图（合并后画图等）。
        if (!hardTool && _isExcelHost) {
            const _rcMerge = RC();
            const MERGE_KW = _rcMerge.MERGE_SHEET_KW;
            const MERGE_TARGET = _rcMerge.MERGE_SHEET_TARGET_KW;
            const MERGE_NON = _rcMerge.MERGE_SHEET_NON_KW;
            // "合并单元格"（可带"取消合并"）→ 属 excel_merge_cells，不在此短路
            const _cellMergeOnly = _rcMerge.MERGE_CELL_ONLY_RE.test(message)
                && !/(工作表|sheet|多表|多张表|数据合并|合并数据)/i.test(message);
            // "合并多个 xlsx/工作簿" → 需文件路径，无法确定性直调
            const _wbMergeOnly = _rcMerge.MERGE_WORKBOOK_ONLY_RE.test(message)
                && !/(工作表|sheet|数据合并|合并数据|合表)/i.test(message);
            if (MERGE_KW.test(m_) && MERGE_TARGET.test(m_) && !MERGE_NON.test(m_) && !_cellMergeOnly && !_wbMergeOnly) {
                // 「合并，并设置成中国红」不能只合并就收工。中国红不在预设里，整句交给模型规划。
                const _mergeStyle = (App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.semanticStyleTarget === 'function')
                    ? (App.ChatV2.turnGoal.semanticStyleTarget(message) || '')
                    : '';
                if (_mergeStyle) {
                    ChatLog.info?.('McpDispatcher', '🎯 合并硬短路放弃：句中还有外观语意「' + _mergeStyle + '」→ 交给模型规划');
                    hostLog('INFO', '合并硬短路放弃：外观语意「' + _mergeStyle + '」需模型规划');
                    return { hit: false, deferToPlanner: true, styleTarget: _mergeStyle, reason: 'semantic-style' };
                }
                hardTool = 'excel_merge_sheets';
                hardArgs = {};
                // 用户在标签栏多选的工作表 → 显式传给后端（未传时后端自动合并除目标表外的所有可见表）
                const _selSheets = (turn && turn.context && turn.context.selectedSheets) || null;
                const _selList = Array.isArray(_selSheets)
                    ? _selSheets.filter(s => typeof s === 'string' && s.trim()) : [];
                if (_selList.length > 0) hardArgs.sourceSheets = JSON.stringify(_selList);
                // "所有/全部/每张" → 强制合并全部可见表（忽略标签选中态）
                if (/(所有|全部|每张|各个|各张)/.test(m_)) hardArgs.mergeAll = true;
                // 显式指定落表名只认带引号的写法（"合并到 '汇总' 表"），避免把"合并到一张表"当成表名
                const _tq = message.match(/合并(?:到|至|成|为)\s*["'「『]\s*([^"'」』]{1,20}?)\s*["'」』]/);
                if (_tq && _tq[1]) hardArgs.targetSheet = _tq[1].trim();
                ChatLog.info?.('McpDispatcher', `🎯 合并工作表硬短路命中 excel_merge_sheets sourceSheets=${hardArgs.sourceSheets || '(后端:除目标表外全部可见表)'} mergeAll=${!!hardArgs.mergeAll} target=${hardArgs.targetSheet || '(默认 合并结果)'}`);
            }
        }

        // ── 拆分工作表硬短路（"按部门拆分成不同Sheet / 按地区拆分"）──
        // 100% 确定性直调 excel_split_by_column，绕开 LLM / agent-loop。
        // 实测（2026-10-09 23:21）：简单拆分被模型自由发挥成"聚合 → 拆分 → 画图"三步
        // （用户吐槽"简单的拆分别搞成聚合拆分"）。与合并硬短路同一范式。
        // 判定 = 拆分动词 + 明确"按X"列名 + 工作表级对象；排除 单元格拆分/拆成文件/复合意图（拆完还要画图等）。
        //
        // 【多任务规划·2026-10-22】复合消息两项任务都要做：
        //   ① "…拆分成不同Sheet，然后分别对工作表美化" → 拆分成功后按产物表逐表确定性美化
        //      （splitThenFormat 链，见下方执行期）；旧实现只跑拆分就收尾，美化被整个吞掉。
        //   ② 句尾还挂着其它任务（排序/筛选/导出/去重/清洗…）→ 不抢戏，放行 agent-loop
        //      多步处理（_splitNon 扩容）；这些任务没有确定性链，交给 LLM 规划才不丢。
        //   ③ 用户点名了具体样式但确定性参数覆盖不全 → 同样放行（uncovered），不硬调半套美化。
        if (!hardTool && _isExcelHost) {
            const _splitVerb = /(拆分|拆开|分开)/.test(m_);
            const _splitTarget = /(工作表|sheet|新表|多张?表)/i.test(m_);
            const _splitNon = /(单元格|文件|工作簿|画图|图表|聚合|统计|透视|饼图|柱状图|折线图|排序|筛选|导出|写入|填写|填充|冻结|去重|清洗|删除|清除|复制|粘贴|追加|汇总|求和|分析|总结|预测|趋势|占比|环比|同比|对比|分组|转换|替换)/i.test(m_);
            const _cm = message.match(/按\s*[「'""]?([^，。,、！？\s]{1,8})[」'""]?\s*(?:进行|来)?(?:拆分|拆开|分开)/);
            if (_splitVerb && _splitTarget && !_splitNon && _cm && _cm[1]) {
                const _splitStyle = (App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.semanticStyleTarget === 'function')
                    ? (App.ChatV2.turnGoal.semanticStyleTarget(message) || '')
                    : '';
                if (_splitStyle) {
                    ChatLog.info?.('McpDispatcher', '🎯 拆分硬短路放弃：句中还有外观语意「' + _splitStyle + '」→ 交给模型规划');
                    hostLog('INFO', '拆分硬短路放弃：外观语意「' + _splitStyle + '」需模型规划');
                    return { hit: false, deferToPlanner: true, styleTarget: _splitStyle, reason: 'semantic-style' };
                }
                const _trailingFmt = detectFormatIntent(message);
                if (_trailingFmt === 'uncovered') {
                    // 复合美化要求无法被确定性参数全覆盖 → 放弃硬短路，放行 agent-loop 多步处理
                    ChatLog.info?.('McpDispatcher', `🎯 拆分+美化硬短路放弃：句尾美化要求覆盖不全 → 放行 agent-loop 多步处理`);
                    hostLog('INFO', '拆分+美化硬短路放弃（多要求排他）：句尾美化要求覆盖不全 → 放行下游');
                } else {
                    hardTool = 'excel_split_by_column';
                    hardArgs = { column: _cm[1].trim(), outputMode: 'sheets' };
                    // 用户明确要求重做/覆盖时覆写同名表；默认 skip，避免二次拆分叠出 _2
                    if (/(重新|再次|重做|覆盖|再拆|重新拆)/.test(m_)) hardArgs.onConflict = 'replace';
                    if (_trailingFmt) {
                        // 复合意图：拆分成功后链式逐表美化。参数已完全确定（column/outputMode/onConflict），
                        // 跳过 LLM 语义补参，避免模型把"美化"字样补进拆分参数里。
                        splitThenFormat = true;
                        splitChainFmt = _trailingFmt.args || {};
                        splitChainLabel = _trailingFmt.label || '商务蓝美化预设';
                        skipRefine = true;
                    }
                    ChatLog.info?.('McpDispatcher', `🎯 拆分工作表硬短路命中 excel_split_by_column column=${hardArgs.column} outputMode=sheets onConflict=${hardArgs.onConflict || 'skip'}（源表未指定→后端取主数据表）${splitThenFormat ? '｜复合意图：成功后逐表美化' : ''}`);
                }
            }
        }

        // ── 批量美化硬短路（"美化工作簿中的所有工作表 / 把所有sheet都美化 / 批量美化"）──
        // 100% 确定性：先枚举可见表（含 usedRange），首表由主调用执行、其余表进链，全程不经过 LLM。
        // 实测（2026-10-22 用户截图）：该请求此前落给 agent-loop，弱模型只调了 1 次
        // excel_format_range（range 还填成 A1，只美化 1 个单元格），随后用"正在依次应用相同预设
        // 至其他工作表，请稍候"的过程语言充数，回合以"任务完成"收场——N 张表只做了 1 张。
        // 与单表"美化表格"（isBareTableBeautify）同一范式，只是目标从 1 张表扩展到整个工作簿。
        if (!hardTool && _isExcelHost && isBatchBeautifyMessage(message)) {
            // 点名了预设盖不住的外观（统一为中国红、改成科技绿…）时，不在这里用关键词把话砍成
            // 「美化全部表 = 商务蓝」。关键词快，但会改了东、丢了西。交给模型按原话规划。
            // 没点名、或点名的就是商务蓝/深蓝时，仍走下面的确定性逐表预设。
            const _styleLeft = (App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.semanticStyleTarget === 'function')
                ? (App.ChatV2.turnGoal.semanticStyleTarget(message) || '')
                : '';
            if (_styleLeft) {
                ChatLog.info?.('McpDispatcher', '🎯 批量美化含未覆盖的外观语意「' + _styleLeft + '」→ 不关键词硬执行，交给模型规划');
                hostLog('INFO', '批量美化放弃硬短路：外观语意「' + _styleLeft + '」需模型规划');
                return { hit: false, deferToPlanner: true, styleTarget: _styleLeft, reason: 'semantic-style' };
            }
            const _batchFmt = detectFormatIntent(message);
            if (_batchFmt === 'uncovered') {
                ChatLog.info?.('McpDispatcher', '🎯 批量美化硬短路放弃：美化要求覆盖不全 → 放行 agent-loop 多步处理');
                hostLog('INFO', '批量美化硬短路放弃（多要求排他）：未全覆盖 → 放行下游');
            } else if (!_batchFmt) {
                // 否定/豁免（"不要美化所有工作表"/"保持格式不变"）→ 整块跳过，绝不硬链上美化
                ChatLog.info?.('McpDispatcher', '🎯 批量美化：消息是否定/豁免语气 → 不批量美化，放行下游');
            } else {
                // 批量链已接管：主调用（首表）会被步骤面板第一行覆盖，不再补裸工具卡
                suppressMainToolCard = true;
                // 枚举可见表（含 usedRange）。⚠ 拿不到清单时【不放行】：实测 2026-10-22 22:04
                // 编辑重发场景，放行后通用兜底把"美化…所有工作表"ngram 匹配到 excel_get_sheets
                //（其 description 含"所有工作表"），执行 1 个工具后把原始 JSON 倾倒进正文冒充完成。
                let _sheets = [];
                let _enumRawHead = '';
                // 【退避重试】实测 2026-10-22 22:10：tryDispatch 紧跟在 gatherContext 的多次桥调用
                // 之后，桥瞬忙时首调可能失败而数秒后同类调用成功（22:04 同会话内先败后成）。
                // 两次尝试 + 350ms 退避，杜绝瞬时竞态把整条确定性链打死。
                for (let _attempt = 1; _attempt <= 2 && _sheets.length === 0; _attempt++) {
                    try {
                        if (_attempt > 1) await new Promise(r => setTimeout(r, 350));
                        const _gotSheets = await callLocalToolWithFallback('excel_get_sheets', '{}');
                        _enumRawHead = String((_gotSheets && _gotSheets.raw) || '').slice(0, 160);
                        _sheets = _parseSheetList(_gotSheets && _gotSheets.raw);
                    } catch (_eBs) {
                        _enumRawHead = 'EXC: ' + ((_eBs && _eBs.message) || _eBs);
                        ChatLog.warn?.('McpDispatcher', '批量美化：excel_get_sheets 枚举异常(第' + _attempt + '次): ' + ((_eBs && _eBs.message) || _eBs));
                    }
                }
                if (_sheets.length === 0) {
                    // 【第三条枚举通道】直连桥的 GetSheets()——与 CallLocalTool 是不同方法，
                    // 桥"个别方法不可调用"（实测的 "e is not a function" 类故障）时仍可能可用。
                    // 只回表名（无 usedRange）也没关系：范围交给降级阶梯逐级解析。
                    try {
                        const _cands = _collectBridgeCandidates();
                        for (const c of _cands) {
                            try {
                                if (c.b && typeof c.b.GetSheets === 'function') {
                                    const _rawG = await c.b.GetSheets();
                                    let _arr = null;
                                    try { _arr = JSON.parse(typeof _rawG === 'string' ? _rawG : JSON.stringify(_rawG)); } catch (_eP) { _arr = null; }
                                    if (Array.isArray(_arr) && _arr.length > 0) {
                                        _sheets = _arr.map(x => (typeof x === 'string')
                                            ? { name: x, usedRange: '' }
                                            : { name: String((x && x.name) || ''), usedRange: String((x && x.usedRange) || '') })
                                            .filter(x => x.name);
                                        if (_sheets.length > 0) {
                                            _enumRawHead = 'GetSheets兜底:' + c.label;
                                            hostLog('INFO', '[批量美化] CallLocalTool 枚举失败，经 ' + c.label + '.GetSheets() 兜底枚举到 ' + _sheets.length + ' 张表');
                                            ChatLog.info?.('McpDispatcher', `[批量美化] GetSheets() 兜底枚举成功（${c.label}）：${_sheets.length} 张表`);
                                            break;
                                        }
                                    }
                                }
                            } catch (_eG) { /* 试下一个候选 */ }
                        }
                    } catch (_eG2) { /* 候选收集失败 → 走报错分支 */ }
                }
                if (_sheets.length > 0) {
                    // 【选区范围判断·检测期逐表解析】get_sheets 的 usedRange 对从未激活过的工作表
                    // 不可信（Excel 维度缓存不刷新 → 返回光标单格如 G30 或空串，实测 2026-10-22）。
                    // 这里就用降级阶梯把每张表的范围定死：主调用也拿可靠范围，解析不出的表
                    // 进 skipped 名单——汇总里点名 + 给原因，绝不静默丢弃、更不拿单格范围硬调。
                    if (!_RR()) {
                        // 共享模块未加载（前端版本不一致/WebView2 缓存了旧 index.html）：
                        // 可靠性判断全部走安全侧。必须大声记日志，否则现场只表现为"静默放行"，
                        // 无法与"工作簿是空表"区分（实测 2026-10-22 21:40 事故）。
                        ChatLog.error?.('McpDispatcher', '⚠️ services/range-resolver.js 未加载！范围可靠性判断失效（前端版本不一致）→ 批量美化将逐表跳过');
                        hostLog('ERROR', 'range-resolver 模块未加载（前端版本不一致/缓存）→ 批量美化逐表跳过；请完全退出 Excel 重开');
                    }
                    const _resolved = [];
                    const _unresolved = [];
                    for (const s of _sheets) {
                        const r = await resolveSheetRangeLadder(s.name, s.usedRange);
                        if (r) _resolved.push({ name: s.name, range: r.range, via: r.via });
                        else _unresolved.push(s.name);
                    }
                    if (_resolved.length > 0) {
                        const _first = _resolved[0];
                        hardTool = 'excel_format_range';
                        hardArgs = _buildChainFormatArgs(_batchFmt && _batchFmt.args, _first.name, _first.range);
                        batchBeautifyRest = _resolved.slice(1);
                        batchBeautifyFmt = (_batchFmt && _batchFmt.args) || {};
                        batchBeautifyLabel = (_batchFmt && _batchFmt.label) || '商务蓝美化预设';
                        batchBeautifySkipped = _unresolved;
                        skipRefine = true;   // 参数已完全确定（逐表 sheet/range + 样式），跳过 LLM 语义补参
                        ChatLog.info?.('McpDispatcher', `🎯 批量美化硬短路命中 excel_format_range：${_sheets.length} 个可见表，可靠 ${_resolved.length} 张（首表=${_first.name}@${_first.range} 主调用，${batchBeautifyRest.length} 张进链）${_unresolved.length ? '，跳过 ' + _unresolved.length + ' 张（' + _unresolved.join('、') + '）' : ''}`);
                        hostLog('INFO', `批量美化硬短路：可靠 ${_resolved.length}/${_sheets.length} 张（首表 ${_first.name} 主调用）${_unresolved.length ? '，跳过=' + _unresolved.join(',') : ''}`);
                    } else {
                        // 【全部不可靠 → 跳过汇报而非放行】实测 2026-10-22 21:40：旧版此处静默放行
                        // agent-loop，弱模型编造"已美化"被零工具守卫道歉，用户看到 gaslight 而非真相
                        // （空工作簿 / range-resolver 模块未加载，两种诱因同一表现）。
                        // 改为确定性产出：不硬调任何格式化（没有可靠范围可调），逐表点名 + 给原因。
                        hardTool = 'excel_format_range';
                        hardArgs = {};
                        batchSkipAll = true;
                        batchBeautifyRest = [];
                        batchBeautifySkipped = _unresolved;
                        batchBeautifyFmt = (_batchFmt && _batchFmt.args) || {};
                        skipRefine = true;
                        ChatLog.warn?.('McpDispatcher', `🎯 批量美化：${_sheets.length} 个可见表全部无法定位数据范围（${_unresolved.join('、')}）→ 产出逐表跳过汇报（不调 format_range、不经过 LLM）`);
                        hostLog('WARN', `批量美化：全部 ${_sheets.length} 张表范围不可靠 → 跳过汇报（${_unresolved.join(',')}）`);
                    }
                } else {
                    // 【枚举失败 → 确定性诚实报错】不再放行：放行后 get_sheets 会被通用兜底
                    // 误选中并倾倒原始 JSON（22:04 事故）。跳过主调用，产出明确报错。
                    hardTool = 'excel_format_range';
                    hardArgs = {};
                    batchSkipAll = true;
                    batchEnumFail = true;
                    batchEnumDiag = String(_enumRawHead || '').slice(0, 100);
                    batchBeautifyRest = [];
                    skipRefine = true;
                    ChatLog.error?.('McpDispatcher', '🎯 批量美化：枚举不到可见工作表 → 诚实报错（不放行）。raw头=' + _enumRawHead);
                    hostLog('ERROR', '批量美化：excel_get_sheets 枚举失败/解析为空 → 诚实报错（不放行）。raw头=' + _enumRawHead);
                }
            }
        }

        // ── 分组：按列或按时间粒度汇总，再交给后面的功能（图表 / 落表）──
        // PPT 已在上面吃掉分组后的行。这里处理「只分组」以及「分组后再画图」。
        // 分组结果写入「分组结果」表，源表不动。
        if (!hardTool && _isExcelHost) {
            const plan = (turn && turn.context && turn.context.groupPlan) || parseGroupPlan(message, []);
            const canGroup = plan && plan.collapse && plan.follow !== 'ppt' && (plan.grain ? plan.dateColumn : plan.dimColumn);
            if (canGroup) {
                const targetSheet = '分组结果';
                if (plan.grain && plan.dateColumn) {
                    hardTool = 'excel_date_group';
                    hardArgs = { dateColumn: plan.dateColumn, groupBy: plan.grain, targetSheet: targetSheet };
                    if (plan.valueColumn) hardArgs.aggColumn = plan.valueColumn;
                } else {
                    hardTool = 'excel_aggregate';
                    hardArgs = { dimColumn: plan.dimColumn, aggFunc: plan.aggFunc || 'sum', targetSheet: targetSheet };
                    if (plan.valueColumn) hardArgs.valueColumn = plan.valueColumn;
                }
                if (plan.sheet) hardArgs.sheet = plan.sheet;
                if (plan.follow === 'chart') {
                    hardAfter = {
                        tool: 'excel_create_chart',
                        fromGroup: true,
                        rangeWidth: plan.grain ? 3 : 2,
                        args: { chartType: plan.chartType || 'auto', title: plan.label || '分组', sheet: targetSheet }
                    };
                }
                ChatLog.info?.('McpDispatcher', `🎯 分组硬短路 ${hardTool} label=${plan.label} follow=${plan.follow || '仅分组'}`);
            }
        }

        // 链式组合（computer_ 多工具："打开→播放"）：进入链前先注入现代步骤面板骨架并逐步填充，
        // 避免执行期（尤其含两次 OCR 验证的第二步，耗时数秒）气泡一直停"三个点"、操作已完成却无进度可看。
        // 与 skill/text 多任务一致的"骨架先出→执行中填色"体验。
        if (hardTool && !skipRefine && !(hardArgs && hardArgs.deleteMatchingRows) && !(hardArgs && hardArgs.beautifyPreset)) {
            hardArgs = await refineArgsFromSemantics(hardTool, message, hardArgs, turn, opts && opts.signal, opts && opts.toolRegistry);
            if (hardAfter && hardAfter.args) {
                hardAfter.args = await refineArgsFromSemantics(hardAfter.tool, message, hardAfter.args, turn, opts && opts.signal, opts && opts.toolRegistry);
            }
        }

        let chainSteps = null;
        const chainPatch = (statusArr) => {
            try {
                if (!sessionStore || !sessionStore.updateTurn || !turn || !turn.id) return;
                const render = window.App?.ChatV2?.renderers?.stepsPlan?.renderStepsPlanHtml;
                if (!render || !Array.isArray(chainSteps)) return;
                const arr = chainSteps.map((s, i) => Object.assign({}, s, { status: statusArr[i] || s.status }));
                sessionStore.updateTurn(turn.id, { _enableSteps: true, _steps: arr, _planHtml: render(arr) });
            } catch (e) { ChatLog.warn?.('McpDispatcher', '链式面板注入失败(不阻断执行): ' + (e?.message || e)); }
        };

        // 【执行登记收集】硬短路直调不经 agent-loop 的工具循环 → turn 的执行记录里没有这些调用
        // → turn-goal/覆盖度层判"format 未完成"，追加一轮空参 format_range（1ms 失败，
        // 用户看到莫名的 ❌ 步骤，实测 2026-10-22 22:16）。链内每次成功调用都收集到这里，
        // 最终 return 前统一登记进 turn._executedToolCalls/_executedTools。
        const _hardExecCalls = [];

        /**
         * 【多任务规划·逐表美化链】对 targets 里的每张表确定性美化，实时驱动步骤面板。
         * 拆分后美化 / 批量美化共用：targets=[{name, range?}]，range 缺省时先 get_used_range 补
         * （空 range 直调会走后端"当前选区"兜底做错表）。单表失败只标记该表并继续，最后返回
         * 真实成败名单——调用方负责写 steps 注记，失败必须点名，不得伪造"全部完成"。
         * @returns {Promise<{ok:string[], fail:Array<{name:string, reason:string}>, ranges:Object}>}
         */
        const runBeautifyChain = async (targets, fmtArgs) => {
            const okNames = [];
            const failures = [];
            const ranges = {};
            // 链一开跑就有步骤面板可看 → 主调用不再补裸工具卡（避免"excel_format_range {} · 表名"
            // 那种看不懂的重复行，实测 2026-10-22 22:31）
            if (Array.isArray(targets) && targets.length > 0) suppressMainToolCard = true;
            // 批量美化时首行就是主调用本身 → 直接用表名命名，别再叫泛泛的"美化表格"
            const _firstStep = stepMeta(hardTool, hardArgs);
            if (hardTool === 'excel_format_range' && hardArgs && hardArgs.sheet) {
                _firstStep.title = '美化「' + String(hardArgs.sheet) + '」';
                _firstStep.detail = 'excel_format_range ' + JSON.stringify({
                    sheet: String(hardArgs.sheet),
                    range: hardArgs.range || '(UsedRange)'
                });
            }
            chainSteps = [{ ..._firstStep, status: 'done' }];
            targets.forEach(t => {
                chainSteps.push({
                    title: '美化「' + t.name + '」',
                    detail: 'excel_format_range ' + JSON.stringify({ sheet: t.name }),
                    status: 'pending'
                });
            });
            chainPatch(chainSteps.map(s => s.status));
            for (let i = 0; i < targets.length; i++) {
                const t = targets[i];
                chainSteps[i + 1].status = 'running';
                chainPatch(chainSteps.map(s => s.status));
                let done = false;
                let reason = '';
                try {
                    // 【选区范围判断 + 降级】种子范围不可靠（单格/空——未激活表的陈旧 UsedRange，
                    // 实测拆分新表返回 G30）→ 逐级升级解析；三级都不可靠 → 跳过该表并说明原因，
                    // 绝不拿单格范围硬调（那只会在 1 个格子上"假成功"），也不空 range 直调
                    //（后端会用当前选区兜底，把用户停在别的表上的选区格式化掉）。
                    const resolved = await resolveSheetRangeLadder(t.name, t.range || '');
                    if (!resolved) {
                        reason = '未能定位数据范围（该表可能未在 Excel 中打开过或没有数据），已跳过';
                    } else {
                        if (resolved.via !== 'get_sheets' && (t.range || '') !== resolved.range) {
                            ChatLog.info?.('McpDispatcher', `🔁 范围解析降级: ${t.name} 种子="${t.range || '(空)'}" 不可靠 → 经 ${resolved.via} 解析为 ${resolved.range}`);
                            hostLog('INFO', `范围解析降级: ${t.name} → ${resolved.range}（经 ${resolved.via}）`);
                        }
                        ranges[t.name] = resolved.range;
                        const args = _buildChainFormatArgs(fmtArgs, t.name, resolved.range);
                        const rFmt = await callLocalToolWithFallback('excel_format_range', JSON.stringify(args));
                        let pFmt = null;
                        try { pFmt = typeof rFmt.raw === 'string' ? JSON.parse(rFmt.raw) : rFmt.raw; } catch (_eP) { pFmt = rFmt.raw; }
                        // 成败只看信封（success/isError），不看结果文本：后端部分成功时会在文本里
                        // 追加"表头美化失败（已跳过）"这类警告，用文本正则判失败会误杀。
                        const fmtOk = !!(pFmt && (pFmt.success === true || pFmt.isError === false));
                        if (fmtOk) {
                            done = true;
                            _hardExecCalls.push({ name: 'excel_format_range', args: args });
                        }
                        else reason = (pFmt && (pFmt.error || (pFmt.isError && pFmt.result))) || '后端返回失败';
                    }
                } catch (_eF) { reason = (_eF && _eF.message) || '调用异常'; }
                chainSteps[i + 1].status = done ? 'done' : 'failed';
                if (done) okNames.push(t.name);
                else failures.push({ name: t.name, reason: String(reason || '未知原因').slice(0, 120) });
                chainPatch(chainSteps.map(s => s.status));
            }
            return { ok: okNames, fail: failures, ranges: ranges };
        };

        /** 把逐表美化结果写成 steps 注记（orchestrator 产物摘要会渲染 steps）+ 结构化字段（验收/测试） */
        const writeBeautifyVerdict = (parsedObj, chainResult, subject, styleLabel) => {
            try {
                if (!parsedObj || typeof parsedObj !== 'object') return;
                const failNames = chainResult.fail.map(f => f.name);
                const label = styleLabel || '商务蓝美化预设';
                const note = chainResult.fail.length === 0
                    ? '已对' + subject + ' ' + chainResult.ok.length + ' 个工作表应用' + label
                    : ('已美化' + subject + ' ' + chainResult.ok.length + ' 个工作表（' + label + '）；失败 ' + chainResult.fail.length + ' 个（'
                        + chainResult.fail.map(f => f.name + (f.reason ? ':' + f.reason : '')).join('、') + '）');
                if (!Array.isArray(parsedObj.steps)) parsedObj.steps = [];
                parsedObj.steps = parsedObj.steps.concat([note]);
                parsedObj.beautified = chainResult.ok;
                if (chainResult.fail.length > 0) parsedObj.beautifyFailed = failNames;
            } catch (_eNote) { /* 注记失败不影响主结果 */ }
        };

        if (hardTool && hardAfter) {
            chainSteps = [
                { ...stepMeta(hardTool, hardArgs), status: 'running' },
                { ...stepMeta(hardAfter.tool, hardAfter.args), status: 'pending' }
            ];
            chainPatch(['running', 'pending']);   // 立即出骨架面板，替代"三个点"
        }

        if (hardTool && hardArgs && hardArgs.deleteMatchingRows) {
            try {
                const removed = await _deleteRowsByValue(hardArgs.matchValue, hardArgs.sheet || '');
                const ok = removed && removed.success !== false;
                return {
                    hit: true,
                    hardMatched: true,
                    toolName: 'excel_delete_row',
                    args: hardArgs,
                    score: 999,
                    execResult: _mkExecResult({
                        ok: ok,
                        result: removed,
                        error: ok ? null : (removed && removed.error) || '删除失败'
                    })
                };
            } catch (eDel) {
                return { hit: true, hardMatched: true, toolName: 'excel_delete_row', args: hardArgs, score: 999, error: friendlyBridgeError(eDel) };
            }
        }

        if (hardTool) {
            console.log('[McpDispatcher] 🎯 关键词硬短路命中: ' + hardTool + ' msg=' + message);
            try {
                // 【修复】多来源回退调用 CallLocalTool：
                // 某些宿主环境下 window.BridgeAccessor.getBridge() 返回的包装桥对个别方法不可调用
                // （报 "e is not a function"），而裸宿主对象 App.excelBridge / window.excelBridge 可用
                // （hostLog 正走裸桥且一直正常）。此处依次尝试所有可用来源，取第一个成功的。
                try {
                    const _labels = _collectBridgeCandidates().map(function (c) { return c.label + ':' + (typeof (c.b && c.b.CallLocalTool)); });
                    hostLog('INFO', `[diag] hardTool=${hardTool} hasBA=${!!window.BridgeAccessor} cands=${_labels.join(',')}`);
                } catch (_de) { hostLog('WARN', '[diag] 桥信息记录失败: ' + (_de && _de.message)); }
                let raw = null;
                // hostObjects defaultSyncMethod='async'：CallLocalTool 返回 Promise，必须 await 拿真实结果
                if (batchSkipAll) {
                    // 【批量美化·全跳过】全部表都定位不到可靠数据范围 → 不调用 format_range
                    // （没有范围可调，调了就是单格假成功或做错表），合成成功信封走汇报路径，
                    // 由下方批量美化块产出"逐表点名跳过"的诚实汇总。
                    raw = JSON.stringify({ success: true, isError: false, batchSkipAll: true });
                } else {
                    try {
                        const _callRes = await callLocalToolWithFallback(hardTool, JSON.stringify(hardArgs));
                        raw = _callRes.raw;
                    } catch (callErr) {
                        console.warn('[McpDispatcher] ❌ 所有桥来源 CallLocalTool 均不可用:', callErr && callErr.message);
                        hostLog('ERROR', `硬短路工具 ${hardTool} 执行失败（所有桥来源均不可用）: ${callErr && callErr.message}`);
                        return { hit: true, hardMatched: true, toolName: hardTool, args: hardArgs || {}, score: 999, error: friendlyBridgeError(callErr) };
                    }
                }
                let parsed = null;
                try { parsed = typeof raw === 'string' ? JSON.parse(raw) : raw; } catch { parsed = raw; }
                const ok = !!(parsed?.success === true || parsed?.isError === false);
                console.log('[McpDispatcher] 后端 CallLocalTool 返回:', parsed);
                if (markDupRows && ok && parsed && typeof parsed === 'object') {
                    try { await _highlightDuplicateRows(parsed, hardArgs); }
                    catch (eMark) { ChatLog.warn?.('McpDispatcher', '标记重复行失败: ' + ((eMark && eMark.message) || eMark)); }
                }
                if (hardArgs && hardArgs.markLiteral) {
                    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
                        parsed = { success: ok, result: parsed, markLiteral: hardArgs.markLiteral, markColorName: hardArgs.markColorName || '黄色' };
                    } else {
                        parsed.markLiteral = hardArgs.markLiteral;
                        parsed.markColorName = hardArgs.markColorName || '黄色';
                    }
                }
                // 链式：组合意图"打开并播放" → 打开成功后追加播放键。
                // 打开"成功"判定：既要无异常(ok)，返回结果也不得含失败标记(not_found/未检测到等)——
                // 否则即使 finally 返回 ok，也不能标"完成"，杜绝"没打开却显示打开成功"这类假确认。
                const step1Fail = !ok || (() => {
                    const o = parsed && typeof parsed === 'object' ? parsed : {};
                    const bits = [o.result, o.message, o.resultText, o.error, o.err].filter(v => typeof v === 'string').join(' ');
                    return o.success === false || o.isError === true ||
                        /not_found|未找到|未检测|不存在|找不到|失败|error|exception|无法|assert_failed/i.test(bits);
                })();
                let chainNote = '';
                if (hardAfter) {
                    const withSkeleton = () => {
                        if (!chainSteps) chainSteps = [
                            { ...stepMeta(hardTool, hardArgs), status: 'done' },
                            { ...stepMeta(hardAfter.tool, hardAfter.args), status: 'running' }
                        ];
                    };
                    if (step1Fail) {
                        // 第一步失败（打开迁移/切换窗口失败）：标红，跳过第二步，不伪造"完成"
                        withSkeleton();
                        chainSteps[0].status = 'failed';
                        chainSteps[1].status = 'pending';
                        chainPatch(['failed', 'pending']);
                        chainNote = '，第一步未完成';
                    } else {
                        withSkeleton();
                        chainSteps[0].status = 'done';
                        chainSteps[1].status = 'running';
                        chainPatch(['done', 'running']);
                        try {
                            if (hardAfter.fromGroup) {
                                const sheet = (parsed && (parsed.writtenTo || parsed.targetSheet)) || (hardArgs && hardArgs.targetSheet) || '分组结果';
                                let n = 0;
                                if (parsed && Array.isArray(parsed.data)) n = parsed.data.length;
                                else if (parsed && typeof parsed.groupCount === 'number') n = parsed.groupCount;
                                else if (parsed && Array.isArray(parsed.rows)) n = Math.max(0, parsed.rows.length - 1);
                                const width = hardAfter.rangeWidth === 3 ? 'C' : 'B';
                                hardAfter.args.dataRange = sheet + '!A1:' + width + Math.max(2, n + 1);
                                hardAfter.args.sheet = sheet;
                            }
                            const _r2 = await callLocalToolWithFallback(hardAfter.tool, JSON.stringify(hardAfter.args));
                            const r2 = _r2.raw;
                            let p2 = null;
                            try { p2 = typeof r2 === 'string' ? JSON.parse(r2) : r2; } catch { p2 = r2; }
                            // 第二步成败：success/isError 之外，结果含失败标记即判失败（避免假确认）。
                            const p2t = String(p2?.result ?? '');
                            let failed = (p2?.success === false || p2?.isError === true) || /^not_found|^not_playing|未检测到|未找到|不存在|失败/.test(p2t);
                            let note = '';
                            if (hardAfter.tool === 'computer_play_media') {
                                // 三级：已有播/触发/失败，避免"已在播却误标红✕"的假阴性
                                if (/验证正在播放/.test(p2t)) note = '，已验证正在播放';
                                else if (failed) note = '，未能确认正在播放';
                                else note = '，已触发播放';
                            } else if (hardAfter.tool === 'computer_capture_window') {
                                // 截图结果即交付物：把干净的 path 直接作为链的 result，供 orchestrator 渲染成图片（不追加文案污染路径）
                                if (!failed) parsed = { ...(typeof parsed === 'object' ? parsed : {}), result: p2t };
                                else note = '，未找到该窗口';
                            } else {
                                if (failed) note = '，执行失败';
                            }
                            chainNote = note;
                            chainSteps[1].status = failed ? 'failed' : 'done';
                            console.log('[McpDispatcher] 链式调用:', hardAfter.tool, p2);
                        } catch (e2) {
                            chainNote = '，步骤执行异常';
                            chainSteps[1].status = 'failed';
                            console.error('[McpDispatcher] 链式第二步异常:', e2?.message || e2);
                        }
                    }
                    if (chainNote && parsed && typeof parsed === 'object') {
                        parsed = { ...parsed, result: (parsed.result || '') + chainNote };
                    }
                    // 第二步完成（或第一步失败）→ 实时落定面板最终状态（完成/失败）
                    chainPatch([chainSteps[0].status, chainSteps[1].status]);
                }

                // ── 【多任务规划·拆分后美化】首工具（excel_split_by_column）成功后，
                //    按产物表逐表确定性美化："按销售员拆分成不同Sheet，然后分别对工作表美化"。
                //    旧实现到此就 return，句尾的美化任务被整个吞掉（用户截图：只执行 1 个工具即"任务完成"）。
                //    逐表两步（get_used_range 取显式范围 → format_range 商务蓝预设）全部在
                //    runBeautifyChain 内确定性直调，失败点名、不伪造"全部完成"。
                if (splitThenFormat && !step1Fail && parsed && typeof parsed === 'object') {
                    const _created = Array.isArray(parsed.created) ? parsed.created.filter(s => typeof s === 'string' && s.trim()) : [];
                    const _replaced = Array.isArray(parsed.replaced) ? parsed.replaced.filter(s => typeof s === 'string' && s.trim()) : [];
                    const _targets = _created.length > 0 ? _created : _replaced;
                    if (_targets.length > 0) {
                        const _chainR = await runBeautifyChain(_targets.map(name => ({ name: name })), splitChainFmt || {});
                        writeBeautifyVerdict(parsed, _chainR, '新工作表', splitChainLabel);
                        ChatLog.info?.('McpDispatcher', `🎯 拆分后逐表美化：成功 ${_chainR.ok.length}/${_targets.length}${_chainR.fail.length ? ' 失败=' + _chainR.fail.map(f => f.name).join(',') : ''}`);
                        hostLog('INFO', `拆分后逐表美化：成功 ${_chainR.ok.length}/${_targets.length}${_chainR.fail.length ? ' 失败=' + _chainR.fail.map(f => f.name).join(',') : ''}`);
                    } else {
                        ChatLog.info?.('McpDispatcher', '🎯 拆分后无新建/覆写工作表（可能全部跳过）→ 跳过逐表美化链');
                    }
                }

                // ── 【多任务规划·批量美化】首表（主调用）成功后，其余可见表进同一条美化链 ──
                // "美化工作簿中的所有工作表"：首表已由上面的主调用真实执行（hardArgs 即其参数），
                // 这里把枚举出的其余表逐表美化。成败判定刻意用 ok（信封）而非 step1Fail（文本正则）：
                // excel_format_range 部分成功时结果文本里会有"表头美化失败（已跳过）"这类警告，
                // 用 /失败/ 判会把成功的结果误杀成 step1Fail。
                if (batchBeautifyRest && ok && parsed && typeof parsed === 'object') {
                    if (batchSkipAll) {
                        // 【全跳过汇报】没有任何一张表定位到可靠数据范围：不美化任何东西，
                        // 如实告知逐表检查结果与原因——绝不把这种请求丢给 LLM 编造"已美化"
                        //（实测 2026-10-22 21:40：空表/模块缺失时旧版放行 → 零工具守卫道歉）。
                        if (batchEnumFail) {
                            // 枚举失败（桥未响应/返回异常）→ 明确报错，绝不能让通用兜底把
                            // get_sheets 的原始 JSON 当回答倾倒（22:04 事故）。
                            // 诊断头随报错透出：用户截图即可定位（该会话 hostLog 不落盘，实测 22:10）
                            parsed = {
                                success: false,
                                error: '未能枚举工作簿的工作表清单（Excel 桥未响应或返回异常'
                                    + (batchEnumDiag ? '：' + batchEnumDiag : '') + '），未做任何修改。请重试一次；若反复出现，请完全退出 Excel 后重开再试',
                                steps: []
                            };
                        } else parsed = {
                            success: true,
                            summary: '已逐表检查 ' + batchBeautifySkipped.length + ' 个工作表（'
                                + batchBeautifySkipped.join('、') + '），均未能定位到可美化的数据范围',
                            beautified: [],
                            beautifySkipped: batchBeautifySkipped.slice(),
                            steps: ['已检查：' + batchBeautifySkipped.join('、')
                                + '——均未能定位数据范围（可能为空表，或数据区尚未被 Excel 识别）。'
                                + '若确认表中有数据，请点击这些工作表后重试；若工作簿本就无数据，可先让我写入或导入数据']
                        };
                        ChatLog.info?.('McpDispatcher', '🎯 批量美化：全跳过汇报（' + batchBeautifySkipped.length + ' 张表无可靠范围）');
                        hostLog('INFO', '批量美化：全跳过汇报（' + batchBeautifySkipped.length + ' 张）');
                    } else if (batchBeautifyRest.length > 0) {
                        const _batchR = await runBeautifyChain(batchBeautifyRest, batchBeautifyFmt);
                        // 汇总改写为覆盖全部表的合成结果（orchestrator 的 excelFriendlyText 走 summary+steps 渲染）
                        const _firstName = String(hardArgs.sheet || '');
                        const _allOk = (_firstName ? [_firstName] : []).concat(_batchR.ok);
                        const _allFail = _batchR.fail.map(f => f.name);
                        const _skipBit = batchBeautifySkipped.length > 0
                            ? ('；跳过 ' + batchBeautifySkipped.length + ' 个（' + batchBeautifySkipped.join('、') + '，未能定位数据范围）')
                            : '';
                        const _synth = {
                            success: true,
                            summary: (_allFail.length === 0
                                ? ('已美化 ' + _allOk.length + ' 个工作表：' + _allOk.join('、'))
                                : ('已美化 ' + _allOk.length + ' 个工作表（' + _allOk.join('、') + '）；失败 ' + _allFail.length + ' 个（' + _allFail.join('、') + '）'))
                                + _skipBit,
                            beautified: _allOk,
                            steps: []
                        };
                        if (_allFail.length > 0) _synth.beautifyFailed = _allFail;
                        if (batchBeautifySkipped.length > 0) _synth.beautifySkipped = batchBeautifySkipped.slice();
                        writeBeautifyVerdict(_synth, _batchR, '其余工作表', batchBeautifyLabel);
                        // writeBeautifyVerdict 只填链内表，这里补上首表（主调用已真实执行成功才有 _firstName）
                        _synth.beautified = _allOk;
                        if (_allFail.length > 0) _synth.beautifyFailed = _allFail;
                        if (batchBeautifySkipped.length > 0) {
                            _synth.steps = _synth.steps.concat(['已跳过 ' + batchBeautifySkipped.length + ' 个工作表（未能定位数据范围，可能未在 Excel 中打开过）：'
                                + batchBeautifySkipped.join('、') + '——请点击这些工作表后重试']);
                        }
                        parsed = _synth;
                        ChatLog.info?.('McpDispatcher', `🎯 批量美化：成功 ${_allOk.length}（含首表），失败 ${_allFail.length}${_allFail.length ? '=' + _allFail.join(',') : ''}${batchBeautifySkipped.length ? '，跳过=' + batchBeautifySkipped.join(',') : ''}`);
                        hostLog('INFO', `批量美化：成功 ${_allOk.length}（含首表 ${_firstName || '?'}），失败 ${_allFail.length}${batchBeautifySkipped.length ? '，跳过=' + batchBeautifySkipped.join(',') : ''}`);
                    } else {
                        // 工作簿只有 1 张可见表（或只有 1 张范围可靠）：主调用即全部，补一个可验收的美化名单
                        if (hardArgs.sheet) {
                            parsed.beautified = [String(hardArgs.sheet)];
                            if (!Array.isArray(parsed.steps)) parsed.steps = [];
                            parsed.steps = parsed.steps.concat(['已对 1 个工作表应用' + batchBeautifyLabel]);
                        }
                        // 有跳过表要点名时，原始结果（后端文本、无 summary 字段）补一个可读汇总，
                        // 否则"跳过谁、为什么"在正文里无处安放；后端明细原样收进 steps 不丢信息。
                        if (batchBeautifySkipped.length > 0 && typeof parsed.summary !== 'string') {
                            parsed.summary = '已美化 1 个工作表：' + String(hardArgs.sheet || '')
                                + '；跳过 ' + batchBeautifySkipped.length + ' 个（' + batchBeautifySkipped.join('、') + '，未能定位数据范围）';
                            if (typeof parsed.result === 'string' && parsed.result.trim() && Array.isArray(parsed.steps)) {
                                parsed.steps = parsed.steps.concat([parsed.result.trim()]);
                            }
                        }
                    }
                    // 检测期就解析不出范围、且没进链的表：点名 + 原因 + 结构化字段（不静默丢弃）。
                    // batchSkipAll 已在上方产出完整汇报，这里不再重复追加。
                    if (!batchSkipAll && batchBeautifySkipped.length > 0 && parsed && typeof parsed === 'object') {
                        if (!Array.isArray(parsed.steps)) parsed.steps = [];
                        parsed.steps = parsed.steps.concat(['已跳过 ' + batchBeautifySkipped.length + ' 个工作表（未能定位数据范围，可能未在 Excel 中打开过）：'
                            + batchBeautifySkipped.join('、') + '——请点击这些工作表后重试']);
                        parsed.beautifySkipped = batchBeautifySkipped.slice();
                    }
                }
                // error 优先；后端 Err 时错误常在 result 里（isError=true、error 为空）→ 用 result 兜底，避免一律显示"后端调用失败"
                const errText = parsed?.error || (parsed?.isError && parsed?.result ? safeErrorText(parsed.result) : '');
                // 【执行登记】硬短路直调不经 agent-loop 工具循环 → turn 执行记录缺失 →
                // turn-goal/覆盖度层误判"没做完"，追加一轮空参 format_range（1ms 失败，
                // 用户看到莫名 ❌ 步骤，实测 2026-10-22 22:16）。成功的主调用与链内逐表调用如实登记。
                if (ok && turn) {
                    try {
                        if (!Array.isArray(turn._executedTools)) turn._executedTools = [];
                        if (turn._executedTools.indexOf(hardTool) < 0) turn._executedTools.push(hardTool);
                        if (!Array.isArray(turn._executedToolCalls)) turn._executedToolCalls = [];
                        if (hardArgs && typeof hardArgs === 'object' && Object.keys(hardArgs).length > 0) {
                            turn._executedToolCalls.push({ name: hardTool, args: hardArgs });
                        }
                        for (const _rc of _hardExecCalls) turn._executedToolCalls.push(_rc);
                        if (_hardExecCalls.length > 0) {
                            ChatLog.info?.('McpDispatcher', `[执行登记] ${hardTool} 主调用 + ${_hardExecCalls.length} 次链内调用已登记进 turn 执行记录（覆盖度层可见）`);
                        }
                    } catch (_eReg) { ChatLog.warn?.('McpDispatcher', '执行登记失败(不阻断): ' + ((_eReg && _eReg.message) || _eReg)); }
                }
                // PPT 硬短路执行失败 → 不放假成功，放行通用管道（agent-loop）兜底重试
                if (hardTool === 'ppt_generate' && !ok) {
                    ChatLog.warn?.('McpDispatcher', '[PPT 硬短路] 执行失败，放行通用管道兜底');
                    return { hit: false, toolName: hardTool, reason: 'ppt-dispatch-failed', error: (errText || 'PPT 生成失败') };
                }
                return {
                    hit: true,
                    hardMatched: true,
                    toolName: hardTool,
                    args: hardArgs || {},
                    score: 999,
                    execResult: _mkExecResult({
                        ok,
                        result: parsed,
                        error: ok ? null : (errText || '后端调用失败'),
                        steps: chainSteps || undefined,
                        suppressToolCard: suppressMainToolCard
                    })
                };
            } catch (e) {
                console.error('[McpDispatcher] 硬短路执行异常:', e?.message || e);
                // 【诊断】把异常全文（name/message/stack）落盘，定位"桥方法不是函数"的真实来源
                try {
                    hostLog('ERROR', `[diag] 硬短路异常 tool=${hardTool} name=${e && e.name} msg=${e && e.message} stack=${String((e && e.stack) || '').replace(/\s+/g, ' ').slice(0, 400)}`);
                } catch (_le) { if (typeof ChatLog !== 'undefined') ChatLog.warn('MCP', 'MCP 本地桥 fallback 失败:', _le && _le.message); }
                // PPT 异常同样放行兜底，避免"生成失败却无后续"
                if (hardTool === 'ppt_generate') return { hit: false, toolName: hardTool, reason: 'ppt-dispatch-exception', error: friendlyBridgeError(e) };
                // 选中表格格式化等硬短路：桥方法未暴露（旧 DLL / 缓存页面）时给出可执行提示而非裸 "e is not a function"
                return { hit: true, hardMatched: true, toolName: hardTool, args: hardArgs || {}, score: 999, error: friendlyBridgeError(e) };
            }
        }

        // ── 第 1 步：后端 MatchCapability（最准，走 McpCapabilityRegistry.Match）──
        let bestTool = null;
        let bestScore = 0;
        try {
            const bridge = window.BridgeAccessor ? window.BridgeAccessor.getBridge() : (window.excelBridge || (window.chrome?.webview?.hostObjects?.excelBridge));
            if (bridge && typeof bridge.MatchCapability === 'function') {
                // COM 无参版本（去掉 Optional 后不接受第二个参数）
                // 【修复】async 代理下返回 Promise，await 后才是字符串
                const matchJson = await bridge.MatchCapability(message);
                const matchData = typeof matchJson === 'string' ? JSON.parse(matchJson) : matchJson;
                console.log('[McpDispatcher] 后端 MatchCapability 返回:', matchData);
                if (matchData?.success && Array.isArray(matchData.matches)) {
                    for (const m of matchData.matches) {
                        const tn = m.id || '';
                        if (!tn) continue;
                        // 接受所有本地 MCP 前缀工具（computer_/excel_/ppt_/web_/time_），排除外部客户端工具（需走 CallMcpToolAsync 指定 clientName）
                        if (!tn.startsWith('computer_') && !tn.startsWith('excel_') && !tn.startsWith('ppt_') && !tn.startsWith('web_') && !tn.startsWith('time_')) continue;
                        const score = m.score || 0;
                        console.log('  match ' + tn + ' score=' + score);
                        if (score > bestScore) { bestScore = score; bestTool = tn; }
                    }
                }
            }
        } catch (bErr) {
            console.warn('[McpDispatcher] 后端 MatchCapability 异常:', bErr?.message);
        }

        console.log('[McpDispatcher] 🎯 匹配结果: best=' + bestTool + ' score=' + bestScore);

        // ── 确定性 Skill 排他 ──
        // 图表/美化/清洗/CRUD/格式化这类操作有专门的 skill-pipeline（SmartChartWizard 等），
        // 裸 MCP 工具（excel_create_chart 等）只是其内部子步骤，需要 Skill 层先做选区识别、
        // LLM 参数规划再调用。dispatcher 直接硬调裸工具会因缺参失败且错误被吞成"未知原因"。
        // 因此命中这类工具时放行给 orchestrator 的 skill 路由（返回 hit:false）。
        if (bestTool && McpDispatcher._isSkillGuardedTool(bestTool)) {
            ChatLog.info?.('ORCH', `[Dispatcher-SkillGuard] ${bestTool} 属于确定性 Skill 域，放行 skill-pipeline（不硬短路）`);
            return { hit: false, skillGuard: true };
        }

        if (!bestTool || bestScore < 20) return { hit: false };

        // ── 分析/总结意图排他 ──
        // "分析/总结/解读/概括/评估/建议"这类请求，读数据只是中间步骤，最终必须给文字结论。
        // 若被 excel_read_range 确定性短路，用户只会看到一张数据卡片、拿不到任何分析（实测：
        // "分析当前选中的Excel区域，并给出总结。"只回了表格 + 插入到Excel 按钮）。
        // 命中分析意图 → 放行给 LLM 管道（先读后总结），不用裸工具硬短路。
        if (McpDispatcher._isAnalysisIntent(message)) {
            ChatLog.info?.('ORCH', `[Dispatcher-AnalysisGuard] 分析/总结意图命中，放行 LLM 管道（不硬短路 ${bestTool}）`);
            return { hit: false, analysisGuard: true };
        }

        // ── 必填参数缺失：只补槽，不把整句放进自由规划 ──
        // 短路执行只能传 {}（只有 computer_ 目标文本类工具会被 _buildComputerArgs 补参）。
        // 若命中工具在 schema 里声明了必填参数（如 excel_clean_data 的 operations），空参裸调必然失败。
        // 这里只让模型填这几个必填键；填齐就继续执行。填不齐才放行通用管道。
        // 已经能从原话抽出参数的短命令不进这里。
        let shortArgs = McpDispatcher._buildComputerArgs(bestTool, message) || {};
        if (Object.keys(shortArgs).length === 0 && McpDispatcher._hasRequiredParams(opts.toolRegistry, bestTool)) {
            const filled = await McpDispatcher.fillRequiredSlots({
                toolName: bestTool,
                message,
                toolRegistry: opts.toolRegistry,
                signal: opts && opts.signal
            });
            if (filled && Object.keys(filled).length) {
                shortArgs = filled;
                ChatLog.info?.('ORCH', `[Dispatcher-SlotFill] ${bestTool} 只补了必填槽: ${Object.keys(filled).join(',')}`);
            } else {
                ChatLog.info?.('ORCH', `[Dispatcher-RequiredArgsGuard] ${bestTool} 有必填参数而补槽失败，放行通用管道`);
                return { hit: false, requiredArgsGuard: true, slotFillAttempted: true };
            }
        }

        // ── 直接执行：调后端 CallLocalTool ──
        // 【关键修复】此前这里固定传 '{}'，导致 assert_text/wait_text/locate_text 这类
        //     依赖 text 目标词的 computer 工具永远拿不到参数（空串 IndexOf("") 恒命中首个
        //     OCR 元素 → 假阳性）。现在为带目标文本的 computer 工具从用户消息里解析 text。
        let callArgs = shortArgs;
        try {
            const t0 = Date.now();
            callArgs = await refineArgsFromSemantics(bestTool, message, shortArgs, turn, opts && opts.signal, opts && opts.toolRegistry);
            if (bestTool === 'excel_format_range' && callArgs && typeof callArgs === 'object') {
                const confine = App.ChatV2?.toolRegistryUtil?.confineFormatFillToHeader;
                if (typeof confine === 'function') {
                    if (typeof callArgs.format === 'string') {
                        try { callArgs.format = JSON.parse(callArgs.format); } catch (_) { /* 保持原字符串 */ }
                    }
                    const goalTurn = (turn && (turn.userMessage || turn.message)) ? turn : { userMessage: message };
                    const headerFill = confine(callArgs, goalTurn);
                    if (headerFill) {
                        delete callArgs._headerFill;
                        callArgs.__headerFill = headerFill;
                    }
                }
            }
            const _headerFill = callArgs && callArgs.__headerFill;
            if (callArgs && callArgs.__headerFill) delete callArgs.__headerFill;

            // ── PPT 工具：从前端 context.selection 自动注入 headers+rows ──
            // 后端 BridgeRegistry → IExcelDataReader 兜底经常失败（COM 线程/选区状态/UsedRange 空等），
            // 但前端 orchestrator 在进入 dispatcher 时已经把当前 Excel 选区读好了放在 context.selection 里。
            // 直接用现成数据塞进 args，秒级解决"缺少数据来源"问题，不依赖后端任何兜底。
            if (bestTool && bestTool.startsWith && bestTool.startsWith('ppt_')) {
                const needData = !callArgs.headers && !callArgs.rows && !callArgs.rangeAddress;
                if (needData) {
                    const ctxSel = opts?.context?.selection || turn?.context?.selection;
                    if (ctxSel && ctxSel.success !== false && ctxSel.headers?.length > 0 && ctxSel.rows?.length > 0) {
                        const pptSel = pptRowsFrom(turn, ctxSel);
                        const shaped = shapePptArgs(message, pptSel.headers, pptSel.rows, pptSel.sheetName, pptSel.note);
                        callArgs.headers = shaped.headers;
                        callArgs.rows = shaped.rows;
                        callArgs.autoDesign = true;
                        callArgs.reportTitle = shaped.reportTitle;
                        callArgs.requirement = callArgs.requirement || shaped.requirement;
                        callArgs.sourceLabel = shaped.sourceLabel;
                        console.log('[McpDispatcher] ✅ PPT ' + bestTool + ' 自动注入 selection: ' + ctxSel.headers.length + '列 × ' + ctxSel.rows.length + '行');
                    } else {
                        console.warn('[McpDispatcher] ⚠️ PPT ' + bestTool + ' 无可用 selection 数据（ctxSel=' + JSON.stringify(ctxSel)?.substring(0, 100) + '）');
                    }
                }
            }

            let resultJson = null;
            let _bridgeOk = false;
            // 【修复】多来源回退调用，规避个别包装桥对 CallLocalTool 不可调用
            if (_collectBridgeCandidates().length) {
                try {
                    const _r = await callLocalToolWithFallback(bestTool, JSON.stringify(callArgs));
                    resultJson = _r.raw;
                    _bridgeOk = true;
                    if (_headerFill) {
                        try { await callLocalToolWithFallback(bestTool, JSON.stringify(_headerFill)); }
                        catch (fillErr) { hostLog('WARN', `[header-fill] 表头底色补涂失败: ${(fillErr && fillErr.message) || fillErr}`); }
                    }
                } catch (_ce) {
                    hostLog('WARN', `[bridge-fallback] ${bestTool} 桥调用失败，尝试降级 toolRegistry: ${(_ce && _ce.message) || _ce}`);
                }
            }
            if (!_bridgeOk) {
                // 降级：走前端 toolRegistry.invokeTracked（如果有）
                const tr = opts.toolRegistry;
                if (tr) {
                    const execResult = await tr.invokeTracked(turn, bestTool, callArgs, { round: 1, displayName: bestTool, ctx: { turn } });
                    if (_headerFill && execResult && execResult.ok !== false) {
                        try { await tr.invokeTracked(turn, bestTool, _headerFill, { round: 2, displayName: bestTool, ctx: { turn } }); }
                        catch (fillErr) { hostLog('WARN', `[header-fill] 表头底色补涂失败: ${(fillErr && fillErr.message) || fillErr}`); }
                    }
                    return { hit: true, toolName: bestTool, score: bestScore, args: callArgs, execResult: _mkExecResult({ ...execResult, elapsedMs: Date.now() - t0 }) };
                }
                // 无任何可用执行通道：明确报错，避免"假成功"
                return { hit: true, toolName: bestTool, score: bestScore, args: callArgs, error: friendlyBridgeError(new Error('本地工具桥不可用（CallLocalTool 在所有来源上均失败）')) };
            }
            console.log('[McpDispatcher] CallLocalTool 返回:', resultJson);
            let parsed = null;
            try { parsed = typeof resultJson === 'string' ? JSON.parse(resultJson) : resultJson; } catch { parsed = resultJson; }
            return { hit: true, toolName: bestTool, score: bestScore, args: callArgs, execResult: _mkExecResult({ ok: parsed?.success !== false, result: parsed, raw: resultJson, elapsedMs: Date.now() - t0 }) };
        } catch (execErr) {
            console.error('[McpDispatcher] 执行失败:', execErr?.message);
            return { hit: true, toolName: bestTool, score: bestScore, args: callArgs, error: friendlyBridgeError(execErr) };
        }
    },
    /**
     * 为需要精确目标参数的 computer_ 工具从用户消息中提取参数。
     * 仅解析 text 类工具；其余工具返回 {}（交给后端默认值）。
     * 解析不出目标词时返回 null，由调用方放行给 agent-loop 而非硬短路。
     */
    _buildComputerArgs(toolName, message) {
        // excel_split_by_column：唯一必填参数 column 必须在这里抽出来。
        // 短路路径直接 bridge.CallLocalTool(bestTool, args)，完全绕过 toolRegistry，
        // 既拿不到 completeRequiredArgs 补参，也过不了下面那道"必填参数缺失排他"闸门
        // （闸门依赖注册表查 schema，查不到就静默失效）→ 空参裸调被后端拒
        // （实测"把这份表按'销售区域'列拆分成多个工作表"回"拆分列不能为空"）。
        if (toolName === 'excel_split_by_column') {
            return McpDispatcher._buildSplitColumnArgs(message) || {};
        }
        const needText = ['computer_assert_text', 'computer_wait_text', 'computer_locate_text', 'computer_type_text', 'computer_clipboard_write'];
        if (!needText.includes(toolName)) return {};
        // 优先取引号/书名号内容
        let target = null;
        const m1 = message.match(/["“”『』「」『』【】]([^"“”『』「」『』【】]{1,40})["“”『』「」『』【】]/);
        if (m1) target = m1[1].trim();
        if (!target) {
            // 前置识别语（动词短语）→ 提取其后剩余文本作为目标
            const verbs = [
                /校验(?:屏幕)?(?:是否)?(?:出现|存在|显示|消失)(.*)$/,
                /等(?:屏幕)?(?:出现|看到)(.*)$/,
                /等待(?:屏幕)?(?:出现|看到)(.*)$/,
                /定位(?:屏幕上)?(?:的)?文字(.*)$/,
                /找到(?:屏幕(?:上)?)?(?:的)?文字(.*)$/,
                /输入(?:文字|内容)?(.*)$/,
                /(?:输入|键入)(.*)$/,
                /写(?:入|到)剪贴板(.*)$/
            ];
            for (const re of verbs) {
                const mm = message.match(re);
                if (mm) { target = (mm[1] || '').trim(); if (target) break; }
            }
        }
        if (!target) return null;
        // 去掉结尾可能的语气/命令词
        target = target.replace(/[，。！？、:,;；\"“”]+\s*$/, '').trim();
        if (!target) return null;
        return { text: target };
    },
    /**
     * 从用户消息里确定性抽出 excel_split_by_column 的拆分列（column）。
     * 覆盖说法："按'销售区域'列拆分成多个工作表" / 按销售区域列拆分 / 以地区列拆分 /
     * 按部门把当前数据拆分到多个工作表 / 按销售区域拆分。
     * 抽不到（如"按每1000行拆分"，那是行数不是列）返回 null，交回下游路由/LLM 规划。
     * @param {string} message 用户原始消息
     * @returns {{column:string}|null}
     */
    _buildSplitColumnArgs(message) {
        const m = String(message || '');
        if (!m) return null;
        const QUOTES = '"\'“”‘’「」『』【】';
        const Q = `[${QUOTES}]`;
        // 清洗并否决"抓到的不是列名"的情况（数字开头、"…行"结尾、泛指短语）
        const clean = v => {
            const s = String(v || '').trim().replace(new RegExp(`^[${QUOTES}\\s]+|[${QUOTES}\\s]+$`, 'g'), '');
            if (!s || s.length > 20) return null;
            if (/^\d/.test(s) || /行$/.test(s)) return null;
            if (/[把将]|当前|这个|这份|这些|该表/.test(s)) return null;
            return s;
        };
        const PATTERNS = [
            // ① 按'销售区域'列拆分（带引号）
            new RegExp(`按\\s*${Q}\\s*([^${QUOTES}\\n]{1,20}?)\\s*${Q}\\s*列`),
            // ② 按销售区域列拆分（无引号）
            new RegExp(`按\\s*([^\\s，。、；;()（）${QUOTES}]{1,20}?)\\s*列\\s*(?:来)?(?:拆分|拆成|拆为|拆开|分成|分为|分割)`),
            // ③ 以/用销售区域列拆分
            new RegExp(`[以用]\\s*${Q}?\\s*([^${QUOTES}\\s，。、；;()（）]{1,20}?)\\s*${Q}?\\s*列\\s*(?:来)?(?:拆分|拆成|拆开|分成)`),
            // ④ 按部门把当前数据拆分到多个工作表（列名在"把…拆分"之前）
            /按\s*([^\s，。、；;()（）]{1,20}?)\s*把[^，。、；;\n]{0,20}?(?:拆分|拆成|分成|分割)/,
            // ⑤ 按销售区域拆分（没有"列"字）
            /按\s*([^\s，。、；;()（）]{1,20}?)\s*(?:拆分|拆成|拆为|拆开|分成|分为|分割)/
        ];
        for (const re of PATTERNS) {
            const mm = m.match(re);
            if (mm) {
                const col = clean(mm[1]);
                if (col) return { column: col };
            }
        }
        return null;
    },
    /**
     * 判断工具是否属于"确定性 Skill 域"（应由 skill-pipeline 而非 dispatcher 裸调）。
     * 这些操作需要 Skill 层做选区识别 + LLM 意图规划后再组合裸工具，dispatcher 硬调会缺参失败。
     */
    _isSkillGuardedTool(name) {
        if (!name || typeof name !== 'string') return false;
        // 统一走声明式注册表（唯一事实源）：命中确定性 Skill 域的受护工具即放行给 skill-pipeline。
        // 历史硬编码 GUARDED 名单已移除 —— 归属关系收敛到 config/skill-domain.js 的 guardedTools。
        const sd = App.ChatV2?.skillDomain;
        return sd ? sd.isGuardedTool(name) : false;
    },
    /**
     * 是否为"分析/总结"类请求：读数据只是中间步骤，最终交付物是文字结论。
     * 命中则不做确定性硬短路，交给 LLM 管道（先读后总结）。
     * @param {string} message 用户原始消息
     * @returns {boolean}
     */
    _isAnalysisIntent(message) {
        if (!message || typeof message !== 'string') return false;
        return /(分析|总结|小结|解读|概括|归纳|评估|评价|洞察|趋势|规律|原因|为什么|建议|优劣|优缺点|风险|意义|看点|结论|判断|汇报|说明一下)/.test(message);
    },
    /**
     * 工具是否声明了必填参数（决定"空参硬短路"是否安全）。
     * 唯一事实源是工具 schema 的 required：有必填项 → 空参必失败 → 不短路。
     * @param {object} toolRegistry 前端工具注册表（tryDispatch 的 opts.toolRegistry）
     * @param {string} name 工具名
     * @returns {boolean}
     */
    _hasRequiredParams(toolRegistry, name) {
        try {
            const tool = toolRegistry?.get?.(name);
            // 注册表缺失该工具（会话早期/桥竞态）：无法确认"空参可安全执行"→ 保守视为"需参数"，
            // 放行通用管道由 LLM 规划参数，避免 dispatcher 空参裸调后端
            // （实测 excel_apply_filter 无 range → "范围不能为空"）。
            if (!tool) return true;
            const req = App.ChatV2?.toolRegistryUtil?.['_requiredParamKeys']?.(tool);
            return Array.isArray(req) && req.length > 0;
        } catch { return false; }
    },
    /**
     * 只补必填参数，不规划其它步骤。
     * 模型不可用、超时、或仍缺必填键时返回 null，调用方放行下游。
     * @returns {Promise<object|null>}
     */
    async fillRequiredSlots(opts) {
        const toolName = opts && opts.toolName;
        const message = String((opts && opts.message) || '');
        const tool = opts && opts.toolRegistry && opts.toolRegistry.get && opts.toolRegistry.get(toolName);
        const util = App.ChatV2 && App.ChatV2.toolRegistryUtil;
        const required = (util && typeof util._requiredParamKeys === 'function' && tool)
            ? util._requiredParamKeys(tool) : [];
        if (!tool || !required.length) return null;
        const llm = App.ChatV2 && App.ChatV2.llmIo;
        if (!llm || typeof llm.askJson !== 'function') return null;
        if (opts.signal && opts.signal.aborted) return null;
        const schema = tool.inputSchema || tool.schema || {};
        const props = (schema && schema.properties && typeof schema.properties === 'object') ? schema.properties : {};
        const allowed = Object.keys(props).length ? Object.keys(props) : required.slice();
        const desc = required.map(k => {
            const d = props[k] && props[k].description;
            return '- ' + k + (d ? '：' + String(d).slice(0, 80) : '');
        }).join('\n');
        const prompt = '你只补全工具 ' + toolName + ' 的调用参数。不要规划其它步骤，不要换工具，不要写解释。\n'
            + '用户原话：' + message.slice(0, 400) + '\n'
            + '只输出一个 JSON 对象。键只能是：' + allowed.join(', ') + '\n'
            + '必填键：' + required.join(', ') + '\n'
            + '原话里确定不了的键不要编造，省略该键。\n'
            + desc;
        let timer = null;
        try {
            const timeoutP = new Promise(resolve => {
                timer = setTimeout(() => resolve({ json: null, timeout: true }), 8000);
            });
            const res = await Promise.race([
                llm.askJson({
                    messages: [{ role: 'user', content: prompt }],
                    signal: opts.signal,
                    config: (typeof App !== 'undefined' && App.config) || {}
                }),
                timeoutP
            ]);
            const json = res && res.json;
            if (!json || typeof json !== 'object' || Array.isArray(json)) return null;
            const out = {};
            const okVal = (v) => {
                if (v == null) return false;
                if (typeof v === 'string') return v.trim().length > 0;
                if (Array.isArray(v)) return v.length > 0;
                if (typeof v === 'object') return Object.keys(v).length > 0;
                return true;
            };
            for (const k of allowed) {
                if (okVal(json[k])) out[k] = json[k];
            }
            if (required.some(k => !okVal(out[k]))) return null;
            return out;
        } catch (_) {
            return null;
        } finally {
            if (timer) clearTimeout(timer);
        }
    },
    /** 判断是否为本地 MCP 工具（按前缀） */
    _isLocalMcpTool(name) {
        if (!name || typeof name !== 'string') return false;
        // 所有 MCP 工具有明确前缀约定，加新类型只改这里
        const PREFIXES = ['computer_', 'ppt_', 'excel_', 'browser_', 'wps_'];
        return PREFIXES.some(p => name.startsWith(p));
    },
    _mergeSemanticArgs: mergeSemanticArgs,
    _parseQuarterPeriod: parseQuarterPeriod,
    _parseGroupPlan: parseGroupPlan,
    _applyGroupView: applyGroupView,
    _shapePptArgs: shapePptArgs,
    attachGroupView: attachGroupView,
    isBareTableBeautify: isBareTableBeautify
};

// 把任意错误值转成可读文本：函数/Promise/对象不再被 String() 成 "function () { [native code] }"
function safeErrorText(v) {
    if (v == null) return '未知错误';
    if (typeof v === 'string') return v || '未知错误';
    if (typeof v === 'function') return '后端调用返回了函数对象（可能未正确 await 或桥接异常）';
    if (v instanceof Error) return v.message || String(v);
    if (typeof v === 'object') {
        if (typeof v.message === 'string') return v.message;
        try { const s = JSON.stringify(v); return s && s.length <= 500 ? s : (s ? s.slice(0, 500) + '…' : '未知错误'); } catch (e) { return String(v); }
    }
    return String(v);
}

// 把 WebView2 宿主桥的"方法未暴露"类原始 TypeError 翻译成用户可读、可执行的提示，
// 不再把 "e is not a function" / "Object doesn't support this property or method" 之类的
// 内部错误直接甩给用户（这类错误几乎总是旧版插件仍在运行 / WebView2 缓存了旧页面所致）。
function friendlyBridgeError(e) {
    const msg = (e && e.message) ? e.message : (e == null ? '' : String(e));
    if (/is not a function|doesn't support this property|object doesn't support|undefined is not|找不到|未定义/i.test(msg)) {
        return '当前加载的插件未暴露该桥接方法（疑似旧版插件仍在运行，或 WebView2 缓存了旧页面）。请：① 完全退出 PowerPoint/Word 进程后重新打开；② 在「文件 → 选项 → 加载项」中确认本插件已启用且为最新版本；③ 若仍失败，清掉 %LOCALAPPDATA%\\Microsoft\\Edge\\User Data\\WebView2Cache 后重试。';
    }
    return msg || '未知错误';
}

// 暴露到全局
window.App.ChatV2.mcpDispatcher = McpDispatcher;
// 暴露多来源回退的桥调用器，供 agent-loop / app-mcp 等其他调用点复用，统一规避
// "window.BridgeAccessor 包装桥对个别方法不可调用（e is not a function）" 的环境问题。
window.App.ChatV2.callLocalToolWithFallback = callLocalToolWithFallback;
window.App.ChatV2.collectBridgeCandidates = _collectBridgeCandidates;
window.App.mcpDispatcher = McpDispatcher;