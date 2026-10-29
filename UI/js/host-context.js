/* ============================================================================
 * host-context.js — SmartPilot 宿主上下文
 *
 * 同一套 UI 同时服务 Excel / Word / PowerPoint 三个宿主：
 *   面板 URL 形如 index.html?host=word / ?host=powerpoint（缺省 excel）
 * 本脚本最早执行，负责：
 *   1. 解析宿主类型，暴露全局 window.SmartPilotHost
 *   2. 非 Excel 宿主下隐藏 Excel 专属入口（撤销/重做、引用区域、选区状态栏…）
 *   3. 给 body 加 host-word / host-powerpoint 类，便于后续按宿主定制样式
 *
 * Excel 宿主行为完全不变（所有隐藏逻辑仅在非 Excel 下生效）。
 * ==========================================================================*/
(function () {
    var kind = 'excel';
    try {
        var m = /[?&]host=([^&#]+)/.exec(window.location.search || '');
        if (m) {
            var h = decodeURIComponent(m[1] || '').toLowerCase();
            if (h === 'word' || h === 'winword' || h === 'docx') {
                kind = 'word';
            } else if (h === 'powerpoint' || h === 'ppt' || h === 'pptx') {
                kind = 'powerpoint';
            }
        }
    } catch (e) { /* 解析失败按 Excel 处理 */ }

    // 审计要求：缺省按 Excel 处理是静默的。若 Word/PPT 宿主忘记在 WebView2 导航 URL
    // 追加 ?host=word|powerpoint，UI 会静默以 Excel 身份运行并加载 Excel 专属模块，
    // 导致跨宿主错配却无迹可查。这里显式告警，便于定位。
    if (!m) {
        console.warn('[host-context] 未检测到 ?host= 参数，按 Excel 处理。' +
            '若当前实为非 Excel 宿主，请确认 WebView2 导航 URL 已追加 ?host=word|powerpoint，' +
            '否则 Excel 专属功能会被错误加载。');
    }

    var isExcel = (kind === 'excel');
    window.SmartPilotHost = {
        kind: kind,
        isExcel: isExcel,
        isWord: (kind === 'word'),
        isPowerPoint: (kind === 'powerpoint'),
        // 本地 MCP 工具源（与后端 Config\*_tools.json 的 source 一致）
        source: isExcel ? 'excel-local' : (kind === 'word' ? 'word-local' : 'ppt-local'),
        label: isExcel ? 'Excel' : (kind === 'word' ? 'Word' : 'PowerPoint')
    };
    window.IS_EXCEL_HOST = isExcel;

    function hide(el) {
        if (el && el.style) { el.style.display = 'none'; }
    }

    /** 隐藏仅在 Excel 下才有意义的入口 */
    function hideExcelOnlyEntries() {
        if (isExcel) { return; }
        var ids = [
            'undo-btn', 'redo-btn',                 // 撤销/重做（依赖 Excel UndoBridge）
            'reference-area-panel',                 // 引用区域面板容器
            'reference-add-btn-inline',             // 「添加引用区域」按钮
            'reference-add-hint',                   // 「请在 Excel 中选择区域」提示
            'fsb-selection',                        // 固定状态栏的选区显示
            'excel-connection-widget'               // 网页版 Excel 连接控件
        ];
        for (var i = 0; i < ids.length; i++) {
            hide(document.getElementById(ids[i]));
        }
    }

    function apply() {
        try { document.body.classList.add('host-' + kind); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HOST', '设置宿主类名失败:', e && e.message); }
        hideExcelOnlyEntries();
        // 部分节点由后续脚本动态插入（如 excel-connection-widget defer），多拍几次收口
        setTimeout(hideExcelOnlyEntries, 800);
        setTimeout(hideExcelOnlyEntries, 2000);
        setTimeout(hideExcelOnlyEntries, 5000);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', apply);
    } else {
        apply();
    }

    if (window.ChatLog && typeof window.ChatLog.info === 'function') {
        window.ChatLog.info('HOST', '[host-context] 当前宿主: ' + kind);
    }
})();