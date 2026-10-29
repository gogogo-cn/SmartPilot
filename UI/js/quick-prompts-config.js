// 快捷提示词统一配置模块（唯一数据源，替代分散硬编码）
// 数据模型: { version, items: [{ id, group, icon, label, prompt }] }
// 持久化: localStorage（用户自定义覆盖默认配置）
(function () {
    'use strict';

    var STORAGE_KEY = 'smartpilot.quickPrompts.config';
    var LEGACY_KEY = 'smartpilot.sideToolbar.quickPrompts'; // 旧版侧边栏自定义词（一次性迁移）

    // 默认配置（按宿主区分：Excel / Word / PowerPoint）
    // 数据模型: { id, group, icon, label, prompt }

    // ---------------- Excel 默认集 ----------------
    // 2026-10 之前未改过的默认集。localStorage 若仍是这一份，启动时换成下面的 EXCEL_ITEMS。
    var PREVIOUS_EXCEL_ITEMS = [
        { id: 'p1', group: '数据分析', icon: '📊', label: '区域总结', prompt: '分析当前选中的 Excel 区域，并给出总结。' },
        { id: 'p2', group: '数据分析', icon: '📈', label: '描述统计', prompt: '分析选中数据的描述统计' },
        { id: 'p3', group: '图表', icon: '📊', label: '饼图展示销售额', prompt: '使用饼图按部门展示销售额' },
        { id: 'p4', group: '图表', icon: '📉', label: '生成图表', prompt: '对选中区域生成图表' },
        { id: 'p5', group: '数据清洗', icon: '🧹', label: '智能清洗', prompt: '清洗这份数据' },
        { id: 'p6', group: '数据清洗', icon: '🧹', label: '清洗并说明', prompt: '清洗选中的数据并输出修改说明' },
        { id: 'p7', group: '合并拆分', icon: '🔗', label: '合并工作表', prompt: '将选中工作表进行数据合并' },
        { id: 'p8', group: '合并拆分', icon: '📑', label: '按部门拆分Sheet', prompt: '按部门拆分成不同Sheet' },
        { id: 'p9', group: '合并拆分', icon: '✂️', label: 'VBA拆分工作簿', prompt: '帮我写一段 VBA 代码：拆分工作簿' },
        { id: 'p10', group: '导出', icon: '💾', label: '导出为CSV', prompt: '将当前工作表导出为 CSV 并保存到桌面' },
        { id: 'p11', group: '检查', icon: '🔍', label: '检查错误', prompt: '检查当前工作表是否包含错误' },
        { id: 'p12', group: '美化', icon: '🎨', label: '美化表格', prompt: '美化当前选中的表格区域：设置表头底色、边框、对齐方式与列宽，使其清晰美观' }
    ];

    var EXCEL_ITEMS = [
        { id: 'p0', group: '识图入表', icon: '📷', label: '识别图片写入表', prompt: '打开识图入表，把截图、表格、表单或票据写入当前工作表', action: 'ocr' },
        { id: 'p5', group: '数据清洗', icon: '🧹', label: '智能清洗', prompt: '清洗这份数据' },
        { id: 'p8', group: '合并拆分', icon: '📑', label: '按部门拆分Sheet', prompt: '按部门拆分成不同Sheet' },
        { id: 'p4', group: '图表', icon: '📉', label: '生成图表', prompt: '对选中区域生成图表' },
        { id: 'p1', group: '数据分析', icon: '📊', label: '区域总结', prompt: '分析当前选中的 Excel 区域，并给出总结。' },
        { id: 'p2', group: '数据分析', icon: '📈', label: '描述统计', prompt: '分析选中数据的描述统计' },
        { id: 'p3', group: '图表', icon: '📊', label: '饼图展示销售额', prompt: '使用饼图按部门展示销售额' },
        { id: 'p6', group: '数据清洗', icon: '🧹', label: '清洗并说明', prompt: '清洗选中的数据并输出修改说明' },
        { id: 'p7', group: '合并拆分', icon: '🔗', label: '合并工作表', prompt: '将选中工作表进行数据合并' },
        { id: 'p10', group: '导出', icon: '💾', label: '导出为CSV', prompt: '将当前工作表导出为 CSV 并保存到桌面' },
        { id: 'p11', group: '检查', icon: '🔍', label: '检查错误', prompt: '检查当前工作表是否包含错误' },
        { id: 'p12', group: '美化', icon: '🎨', label: '美化表格', prompt: '美化当前选中的表格区域：设置表头底色、边框、对齐方式与列宽，使其清晰美观' }
    ];

    // ---------------- Word 默认集 ----------------
    var WORD_ITEMS = [
        { id: 'w1', group: '文档处理', icon: '📝', label: '总结全文', prompt: '总结当前 Word 文档的主要内容与结构。' },
        { id: 'w2', group: '文档处理', icon: '✨', label: '润色段落', prompt: '润色选中的段落，使表达更通顺、专业。' },
        { id: 'w3', group: '文档处理', icon: '🌐', label: '翻译选中', prompt: '将选中的内容翻译为英文。' },
        { id: 'w4', group: '排版', icon: '📑', label: '生成目录', prompt: '基于标题 1-3 样式为文档插入目录。' },
        { id: 'w5', group: '排版', icon: '📊', label: '插入表格', prompt: '在光标处插入一个表格并填入示例数据。' },
        { id: 'w6', group: '排版', icon: '🅰️', label: '应用标题样式', prompt: '将选中的段落应用为"标题 1"样式。' },
        { id: 'w7', group: '编辑', icon: '🔍', label: '查找替换', prompt: '在全文中把指定词替换为另一个词。' },
        { id: 'w8', group: '编辑', icon: '💡', label: '插入批注', prompt: '为选中的内容添加一条审阅批注。' },
        { id: 'w9', group: '检查', icon: '✔️', label: '检查错别字', prompt: '检查文档中的错别字与不通顺表述并列出。' },
        { id: 'w10', group: '导出', icon: '💾', label: '导出 PDF', prompt: '将当前文档导出为 PDF。' }
    ];

    // ---------------- PowerPoint 默认集 ----------------
    var PPT_ITEMS = [
        { id: 't1', group: '幻灯片', icon: '📋', label: '总结当前页', prompt: '总结当前幻灯片的核心内容。' },
        { id: 't2', group: '幻灯片', icon: '🗂️', label: '生成大纲', prompt: '根据主题生成一份 PPT 大纲（多页）。' },
        { id: 't3', group: '幻灯片', icon: '➕', label: '新增标题页', prompt: '新增一页标题页并写入标题与副标题。' },
        { id: 't4', group: '内容', icon: '📊', label: '生成图表页', prompt: '新增一页并在其中插入图表展示数据。' },
        { id: 't5', group: '内容', icon: '🖼️', label: '插入图片', prompt: '在当前幻灯片插入一张图片。' },
        { id: 't6', group: '美化', icon: '🎨', label: '美化当前页', prompt: '美化当前幻灯片的排版与配色，使其清晰美观。' },
        { id: 't7', group: '美化', icon: '🧱', label: '应用版式', prompt: '为当前幻灯片应用合适的版式。' },
        { id: 't8', group: '演讲', icon: '🎤', label: '生成演讲者备注', prompt: '为当前幻灯片生成演讲者备注。' },
        { id: 't9', group: '检查', icon: '✔️', label: '通读润色', prompt: '通读全部幻灯片文本并润色表述。' },
        { id: 't10', group: '导出', icon: '💾', label: '导出 PDF', prompt: '将演示文稿导出为 PDF。' }
    ];

    // 兼容旧逻辑：DEFAULT_ITEMS 仍指向 Excel 默认集（isUntouchedDefault 用其 p1..p11 id 判断）
    var DEFAULT_ITEMS = EXCEL_ITEMS;

    // v1 默认条目 id（用于判断存储数据是否为从未自定义的旧默认集，是则升级到新默认）
    var LEGACY_DEFAULT_IDS = ['p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7', 'p8', 'p9', 'p10', 'p11'];

    var _items = null;
    var _subscribers = [];

    function uid() {
        return 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
    }

    // 当前宿主类型：excel / word / powerpoint
    function hostKind() {
        var kind = 'excel';
        try {
            var k = window.SmartPilotHost && window.SmartPilotHost.kind;
            if (k === 'word' || k === 'powerpoint' || k === 'excel') kind = k;
        } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('QUICK', '操作失败:', e && e.message); }
        return kind;
    }

    // 根据当前宿主返回对应的默认快捷提示词
    function hostItems() {
        var kind = hostKind();
        if (kind === 'word') return WORD_ITEMS;
        if (kind === 'powerpoint') return PPT_ITEMS;
        return EXCEL_ITEMS;
    }

    // 各宿主独立存储（Excel 沿用旧 key 以兼容既有自定义；Word/PPT 使用独立 key，避免继承旧的 Excel 默认词）
    function storageKey() {
        var kind = hostKind();
        return kind === 'excel' ? STORAGE_KEY : STORAGE_KEY + '.' + kind;
    }

    function defaults() {
        return { version: 1, items: hostItems().map(function (it) { return Object.assign({}, it); }) };
    }

    // 规范化：去空/去重（按 prompt），补齐缺失字段
    function normalize(items) {
        var seen = {};
        return (items || []).filter(function (it) {
            if (!it || typeof it.prompt !== 'string' || !it.prompt.trim()) return false;
            var key = it.prompt.trim();
            if (seen[key]) return false;
            seen[key] = true;
            return true;
        }).map(function (it) {
            var prompt = it.prompt.trim();
            var row = {
                id: (it.id && String(it.id)) || uid(),
                group: (it.group && String(it.group).trim()) || '常用',
                icon: (it.icon && String(it.icon).trim()) || '',
                label: (it.label && String(it.label).trim()) || prompt,
                prompt: prompt
            };
            if (it.action === 'ocr') row.action = 'ocr';
            return row;
        });
    }

    function persist() {
        try { localStorage.setItem(storageKey(), JSON.stringify({ version: 1, items: _items })); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('QUICK', '操作失败:', e && e.message); }
    }

    function emit() {
        _subscribers.forEach(function (cb) { try { cb(_items); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('QUICK', '操作失败:', e && e.message); } });
    }

    // 判断一组条目是否与给定默认集完全一致（id + prompt 全匹配）
    function matchesDefaultSet(cfgItems, defItems) {
        if (!cfgItems || !defItems || cfgItems.length !== defItems.length) return false;
        return defItems.every(function (def) {
            var it = null;
            cfgItems.forEach(function (x) { if (x.id === def.id) it = x; });
            return it && String(it.prompt).trim() === def.prompt;
        });
    }

    // 存储数据是否等于旧默认集（用户从未自定义/增删过）→ 可安全替换为当前宿主的默认集
    // 同时识别：当前 Excel 默认集（p1..p12）与更早的 v1 默认集（p1..p11）
    function isUntouchedDefault(cfg) {
        if (!cfg || !Array.isArray(cfg.items)) return false;
        if (matchesDefaultSet(cfg.items, EXCEL_ITEMS)) return true;
        if (matchesDefaultSet(cfg.items, PREVIOUS_EXCEL_ITEMS)) return true;
        if (cfg.items.length === LEGACY_DEFAULT_IDS.length) {
            return LEGACY_DEFAULT_IDS.every(function (id) {
                var it = null;
                cfg.items.forEach(function (x) { if (x.id === id) it = x; });
                if (!it) return false;
                var def = null;
                PREVIOUS_EXCEL_ITEMS.forEach(function (d) { if (d.id === id) def = d; });
                return def && String(it.prompt).trim() === def.prompt;
            });
        }
        return false;
    }

    function load() {
        if (_items) return _items;
        var items = null;
        var key = storageKey();
        try {
            var raw = localStorage.getItem(key);
            if (raw) {
                var cfg = JSON.parse(raw);
                if (cfg && Array.isArray(cfg.items) && cfg.items.length) {
                    items = isUntouchedDefault(cfg) ? defaults().items : cfg.items;
                }
            }
        } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('QUICK', '操作失败:', e && e.message); }
        if (!items) {
            items = defaults().items;
            // 一次性迁移旧版侧边栏自定义词（去重后追加到「常用」组）
            try {
                var old = localStorage.getItem(LEGACY_KEY);
                if (old) {
                    var arr = JSON.parse(old);
                    if (Array.isArray(arr) && arr.length) {
                        var merged = defaults().items.concat(arr.map(function (t) {
                            return { group: '常用', icon: '', label: String(t).trim(), prompt: String(t).trim() };
                        }));
                        items = normalize(merged);
                    }
                    localStorage.removeItem(LEGACY_KEY);
                }
            } catch (e2) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PROMPTS', '快捷提示词降级也失败:', e2 && e2.message); }
        }
        _items = normalize(items);
        persist();
        return _items;
    }

    function save(items) {
        _items = normalize(items);
        persist();
        emit();
        return _items;
    }

    function getItems() { return load().slice(); }

    function getGroups() {
        var groups = [];
        load().forEach(function (it) { if (groups.indexOf(it.group) === -1) groups.push(it.group); });
        return groups;
    }

    function add(item) {
        var list = load().slice();
        list.push(Object.assign({ id: uid() }, item));
        return save(list);
    }

    function update(id, patch) {
        var list = load().slice();
        var idx = -1;
        list.forEach(function (it, i) { if (it.id === id) idx = i; });
        if (idx === -1) return list;
        list[idx] = Object.assign({}, list[idx], patch, { id: id });
        return save(list);
    }

    function remove(id) {
        return save(load().filter(function (it) { return it.id !== id; }));
    }

    // 全局移动（跨组：移入相邻组的首/尾）
    function move(id, dir) {
        var list = load().slice();
        var idx = -1;
        list.forEach(function (it, i) { if (it.id === id) idx = i; });
        var to = idx + dir;
        if (idx === -1 || to < 0 || to >= list.length) return list;
        var item = list.splice(idx, 1)[0];
        list.splice(to, 0, item);
        return save(list);
    }

    // 组内移动（管理视图 ↑↓ 只在同组内换序，保持其他组不变）
    function moveInGroup(id, dir) {
        var list = load().slice();
        var idx = -1;
        list.forEach(function (it, i) { if (it.id === id) idx = i; });
        if (idx === -1) return list;
        var grp = list[idx].group;
        var idxs = [];
        list.forEach(function (it, i) { if (it.group === grp) idxs.push(i); });
        var pos = idxs.indexOf(idx);
        var toPos = pos + dir;
        if (toPos < 0 || toPos >= idxs.length) return list;
        var item = list.splice(idx, 1)[0];
        // 删除后：目标项在 idxs[toPos] > idx 时位于 idxs[toPos] - 1；
        // 插入到目标项后面（下移）等于 idxs[toPos]，插入到其前面（上移）目标项未偏移也为 idxs[toPos]。
        list.splice(idxs[toPos], 0, item);
        return save(list);
    }

    function reset() {
        _items = defaults().items;
        persist();
        emit();
        return _items;
    }

    function onUpdate(cb) {
        if (typeof cb === 'function') _subscribers.push(cb);
    }

    window.App = window.App || {};
    window.App.QuickPrompts = {
        getItems: getItems,
        getGroups: getGroups,
        defaults: defaults,
        add: add,
        update: update,
        remove: remove,
        move: move,
        moveInGroup: moveInGroup,
        reset: reset,
        save: save,
        onUpdate: onUpdate
    };
})();