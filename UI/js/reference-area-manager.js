﻿﻿/**
 * ============================================================
 * reference-area-manager.js — 引用区域管理器
 * ============================================================
 * 职责：
 *   - 管理当前选区显示（panel.current）
 *   - 管理用户添加的引用区域列表（panel.list）
 *   - 支持多选区（GetMultiSelection）
 *   - 为 AI 执行层提供引用区域上下文（getReferenceRegionsContext）
 *
 * 使用：由 js-app-init.js 调用 init()
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};

    const ChatLog = window.ChatLog || console;

    const ReferenceAreaManager = (function() {
        let currentSelection = null;
        let references = [];
        let initialized = false;
        let isAdding = false;
        // 添加模式下的选区基线：避免把进入监听前的当前选区重复捕获
        let lastPollKey = null;

        const panel = document.getElementById('reference-area-panel');
        const listEl = document.getElementById('reference-list');
        const addBtn = document.getElementById('reference-add-btn');
        const hintEl = document.getElementById('reference-add-hint');
        const fsbSelection = document.getElementById('fsb-selection');

        function formatAddress(ref) {
            if (!ref) return '';
            if (ref.workbookName && ref.sheetName) {
                return '[' + ref.workbookName + ']' + ref.sheetName + '!' + ref.address;
            }
            if (ref.sheetName) {
                return ref.sheetName + '!' + ref.address;
            }
            return ref.address || '';
        }

        function isSameRange(a, b) {
            if (!a || !b) return false;
            return a.address === b.address &&
                   a.sheetName === b.sheetName &&
                   a.workbookName === b.workbookName;
        }

        function selectionKey(sel) {
            return [sel && sel.workbookName || '', sel && sel.sheetName || '', sel && sel.address || ''].join('|');
        }

        async function fetchCurrentSelection() {
            try {
                if (!App.excelBridge) return null;
                const selRaw = await App.excelBridge.GetSelection();
                const sel = (typeof selRaw === 'string') ? JSON.parse(selRaw) : selRaw;
                if (!sel || sel.error || !sel.address) return null;

                let workbookName = sel.workbookName || '';
                if (!workbookName && typeof App.excelBridge.GetActiveWorkbookName === 'function') {
                    try {
                        const wbRaw = await App.excelBridge.GetActiveWorkbookName();
                        workbookName = (typeof wbRaw === 'string') ? wbRaw : (wbRaw?.toString() || '');
                    } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('REF', '引用区域操作失败:', e && e.message); }
                }

                return {
                    address: sel.address,
                    sheetName: sel.sheetName || '',
                    workbookName: workbookName
                };
            } catch (e) {
                ChatLog.warn('ReferenceArea', '获取当前选区失败:', e.message);
                return null;
            }
        }

        function render() {
            if (!panel || !listEl) return;

            if (!currentSelection && references.length === 0) {
                panel.style.display = 'none';
                if (fsbSelection) fsbSelection.textContent = '';
                return;
            }

            panel.style.display = 'block';

            // 同步显示到状态栏
            if (currentSelection && fsbSelection) {
                fsbSelection.textContent = formatAddress(currentSelection);
            } else if (fsbSelection) {
                fsbSelection.textContent = '';
            }

            if (hintEl) {
                hintEl.style.display = isAdding ? 'block' : 'none';
                if (addBtn) {
                    addBtn.textContent = isAdding ? '✕ 取消添加' : '+ 添加';
                    addBtn.disabled = false;
                }
            }

            if (hintEl) {
                hintEl.textContent = isAdding
                    ? '▶ 请在 Excel 中选择一个区域（可跨 Sheet/跨工作簿），选中后自动捕获'
                    : '点击"+ 添加"进入选择模式，然后在 Excel 中选择区域即可自动添加';
            }

            listEl.innerHTML = '';
            if (references.length === 0) {
                listEl.innerHTML = '<div class="reference-empty">暂无引用区域，点击"+ 添加"将 Excel 选区加入引用</div>';
                return;
            }

            references.forEach((ref, index) => {
                const item = document.createElement('div');
                item.className = 'reference-item';
                const label = ref.sheetName ? ref.sheetName : '';
                const wbTag = (ref.workbookName && ref.workbookName !== currentSelection?.workbookName)
                    ? `<span class="reference-wb">[${escapeHtml(ref.workbookName)}]</span>` : '';
                item.innerHTML =
                    '<span class="reference-index">🔗 区域' + (index + 1) + ':</span>' +
                    wbTag +
                    '<span class="reference-address">' + escapeHtml(ref.sheetName ? ref.sheetName + '!' + ref.address : ref.address) + '</span>' +
                    '<button class="reference-remove" data-index="' + index + '" title="删除">×</button>';
                listEl.appendChild(item);
            });

            listEl.querySelectorAll('.reference-remove').forEach(btn => {
                btn.onclick = function() {
                    const idx = parseInt(this.dataset.index, 10);
                    remove(idx);
                };
            });
        }

        async function refresh() {
            currentSelection = await fetchCurrentSelection();
            // 添加模式下，选区发生变化（且为有效多单元格区域）则自动捕获
            if (isAdding && currentSelection && currentSelection.address
                    && !isSingleCell(currentSelection.address)) {
                const key = selectionKey(currentSelection);
                if (key !== lastPollKey) {
                    lastPollKey = key;
                    onSelectionChanged({ type: 'changed', data: currentSelection });
                    return; // onSelectionChanged 内部已 render
                }
            }
            render();
        }

        async function startAdding() {
            if (isAdding) return;

            // 实时从 Excel 获取当前选区
            const liveSelection = await fetchCurrentSelection();

            // 判断是否为"有效选区"：多单元格区域（排除单单元格，因为 Excel 永远有单单元格选区）
            const isValidRange = liveSelection && liveSelection.address && !isSingleCell(liveSelection.address);

            if (isValidRange) {
                const ref = {
                    address: liveSelection.address,
                    sheetName: liveSelection.sheetName || '',
                    workbookName: liveSelection.workbookName || ''
                };

                // 如果当前选区已在引用列表里，说明用户想添加其他区域，进入监听模式
                const alreadyAdded = references.some(r => isSameRange(r, ref));
                if (alreadyAdded) {
                    // 进入监听模式，让用户选择新的区域
                    try {
                        isAdding = true;
                        lastPollKey = selectionKey(liveSelection); // 基线：避免轮询把当前选区重复捕获
                        render();
                        if (App.excelBridge && typeof App.excelBridge.StartSelectionMonitor === 'function') {
                            await App.excelBridge.StartSelectionMonitor();
                        }
                        showToast('当前区域已添加。如需添加其他工作簿的区域，请切换到该工作簿并选择区域，选择完成后自动添加', 'info', 8000);
                    } catch (e) {
                        ChatLog.warn('ReferenceArea', '启动选区监听失败:', e.message);
                        isAdding = false;
                        render();
                        showToast('启动监听失败: ' + e.message, 'error');
                    }
                    return;
                }

                // 新区域，直接添加
                references.push(ref);
                render();
                showToast('已添加引用: ' + formatAddress(ref));
                return;
            }

            // 没有有效选区（单单元格或无选区），进入监听模式等待用户选择
            try {
                isAdding = true;
                lastPollKey = liveSelection ? selectionKey(liveSelection) : null;
                render();
                if (App.excelBridge && typeof App.excelBridge.StartSelectionMonitor === 'function') {
                    await App.excelBridge.StartSelectionMonitor();
                }
                showToast('请在 Excel 中选择区域（可跨 Sheet/跨工作簿）', 'info', 5000);
            } catch (e) {
                ChatLog.warn('ReferenceArea', '启动选区监听失败:', e.message);
                isAdding = false;
                render();
                showToast('启动监听失败: ' + e.message, 'error');
            }
        }

        /**
         * 判断地址是否为单单元格（如 "A1"），单单元格视为无效选区
         */
        function isSingleCell(address) {
            if (!address) return true;
            // 单单元格：不含 ":" 且不含 ","（如 "A1"）
            // 多单元格区域：含 ":"（如 "A1:B3"）或含 ","（如 "A1,B3"）
            return !address.includes(':') && !address.includes(',');
        }

        function stopAdding() {
            isAdding = false;
            render();
            if (App.excelBridge && typeof App.excelBridge.StopSelectionMonitor === 'function') {
                try { App.excelBridge.StopSelectionMonitor(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('REF', '引用区域操作失败:', e && e.message); }
            }
        }

        /**
         * 处理后端推送的选区变化事件
         * 由 VSTO 通过 ExecuteScriptOnWebView 调用:
         *   window.__onExcelSelectionChanged({type: 'changed', data: {address, sheetName, workbookName}})
         */
        function onSelectionChanged(event) {
            if (!event || !event.type) return;
            ChatLog.info('ReferenceArea', '选区事件:', event.type);

            if (event.type === 'started') {
                render();
                return;
            }

            if (event.type === 'stopped') {
                isAdding = false;
                render();
                return;
            }

            if (event.type !== 'changed' || !event.data) return;

            if (!isAdding) {
                // 非添加模式时，只更新当前选区显示
                currentSelection = {
                    address: String(event.data.address || '').replace(/\$/g, '').toUpperCase(),
                    sheetName: event.data.sheetName || '',
                    workbookName: event.data.workbookName || ''
                };
                render();
                return;
            }

            // 添加模式：捕获选区并添加到引用列表
            const ref = {
                address: String(event.data.address || '').replace(/\$/g, '').toUpperCase(),
                sheetName: event.data.sheetName || '',
                workbookName: event.data.workbookName || ''
            };

            if (!ref.address) {
                showToast('未获取到有效选区');
                stopAdding();
                return;
            }

            // 单单元格不算有效引用（Excel 任何时刻都有活动单元格），等待用户拖选区域
            if (isSingleCell(ref.address)) return;

            if (references.some(r => isSameRange(r, ref))) {
                showToast('该区域已添加，请换一个区域');
                // 不停止监听，允许用户继续选其他区域
                return;
            }

            references.push(ref);
            stopAdding();
            render();
            showToast('已添加引用: ' + formatAddress(ref));
        }

        /**
         * 捕获当前 Excel 选区（支持多选区）并添加为引用
         */
        async function captureAndAdd() {
            if (!App.excelBridge) {
                showToast('Excel 桥接不可用');
                return;
            }

            try {
                let areas = [];

                if (typeof App.excelBridge.GetMultiSelection === 'function') {
                    const raw = await App.excelBridge.GetMultiSelection();
                    const data = (typeof raw === 'string') ? JSON.parse(raw) : raw;
                    if (data && Array.isArray(data.areas)) {
                        areas = data.areas;
                    }
                }

                if (areas.length === 0) {
                    const selRaw = await App.excelBridge.GetSelection();
                    const sel = (typeof selRaw === 'string') ? JSON.parse(selRaw) : selRaw;
                    if (sel && sel.address) {
                        areas = [{
                            address: sel.address,
                            sheetName: sel.sheetName || '',
                            workbookName: sel.workbookName || ''
                        }];
                    }
                }

                if (areas.length === 0) {
                    showToast('未获取到有效选区，请先在 Excel 中选择区域');
                    return;
                }

                let addedCount = 0;
                areas.forEach(area => {
                    const ref = {
                        address: String(area.address || '').replace(/\$/g, '').toUpperCase(),
                        sheetName: area.sheetName || '',
                        workbookName: area.workbookName || ''
                    };

                    if (!ref.address) return;

                    if (references.some(r => isSameRange(r, ref))) {
                        showToast('区域 ' + formatAddress(ref) + ' 已存在，跳过');
                        return;
                    }

                    references.push(ref);
                    addedCount++;
                });

                render();

                if (addedCount > 0) {
                    showToast('已添加 ' + addedCount + ' 个引用区域');
                } else {
                    showToast('没有新区域可添加');
                }

            } catch (e) {
                ChatLog.warn('ReferenceArea', '捕获选区失败:', e.message);
                showToast('添加选区失败: ' + e.message);
            }
        }

        function remove(index) {
            if (index >= 0 && index < references.length) {
                references.splice(index, 1);
                render();
            }
        }

        function getReferences() {
            return references.map(r => ({ ...r }));
        }

        /**
         * 获取带完整数据的引用区域上下文
         * 供 AI 执行层（MCP / Skill / Standard）统一使用
         */
        async function getReferenceRegionsContext() {
            const refs = getReferences();
            if (!refs || refs.length === 0) return [];

            const enriched = [];
            for (const ref of refs) {
                const addr = formatAddress(ref);
                try {
                    if (App.excelBridge && typeof App.excelBridge.GetRangeData === 'function') {
                        const refDataRaw = await App.excelBridge.GetRangeData(addr);
                        const refData = (typeof refDataRaw === 'string') ? JSON.parse(refDataRaw) : refDataRaw;
                        if (refData && refData.success && Array.isArray(refData.headers)) {
                            enriched.push({
                                address: addr,
                                sheetName: refData.sheetName || ref.sheetName,
                                workbookName: refData.workbookName || ref.workbookName,
                                headers: refData.headers,
                                rows: refData.rows || [],
                                sampleRows: (refData.rows || []).slice(0, 5),
                                rowCount: refData.rowCount || 0,
                                colCount: refData.colCount || 0
                            });
                            ChatLog.info('ReferenceArea', '读取引用区域成功: ' + addr + ', 行数: ' + (refData.rowCount || 0));
                        } else {
                            ChatLog.warn('ReferenceArea', '读取引用区域数据失败: ' + addr + ' - ' + (refData?.error || '未知错误'));
                        }
                    }
                } catch (e) {
                    ChatLog.warn('ReferenceArea', '获取引用区域上下文失败: ' + addr, e.message);
                }
            }
            ChatLog.info('ReferenceArea', '上下文读取完成: ' + enriched.length + '/' + refs.length);
            return enriched;
        }

        function clear() {
            currentSelection = null;
            references = [];
            render();
        }

        function showToast(msg, type) {
            if (App.showToast) {
                App.showToast(msg, type || 'info');
            } else {
                ChatLog.info('[Toast] ' + msg);
            }
        }

        function escapeHtml(str) {
            if (str == null) return '';
            return String(str).replace(/[&<>"']/g, function(c) {
                return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
            });
        }

        function init() {
            if (initialized) return;

            // 注册 VSTO 回调
            window.__onExcelSelectionChanged = onSelectionChanged;

            // 绑定内联按钮（输入框旁的🔗图标）
            const inlineBtn = document.getElementById('reference-add-btn-inline');
            if (inlineBtn) {
                inlineBtn.addEventListener('click', async function() {
                    if (isAdding) {
                        stopAdding();
                        showToast('已取消添加');
                        return;
                    }
                    await startAdding();
                });
            }

            // 绑定引用面板内的"+ 添加"按钮（兼容旧入口）
            if (addBtn) {
                addBtn.addEventListener('click', async function() {
                    if (isAdding) {
                        stopAdding();
                        showToast('已取消添加');
                        return;
                    }
                    await startAdding();
                });
            }

            initialized = true;
            refresh();
            ChatLog.info('ReferenceAreaManager', '已初始化，选区监听回调已注册');
        }

        return {
            init: init,
            refresh: refresh,
            startAdding: startAdding,
            remove: remove,
            getReferences: getReferences,
            getReferenceRegionsContext: getReferenceRegionsContext,
            clear: clear
        };
    })();

    App.ReferenceAreaManager = ReferenceAreaManager;
    ChatLog.info('ReferenceAreaManager', '已加载');

})();