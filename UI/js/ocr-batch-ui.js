/**
 * ============================================================
 * ocr-batch-ui.js — OCR Batch UI
 * ============================================================
 * 功能：OCR 批量处理界面，提供批量图片上传、识别进度展示、结果预览与导出。
 *
 * 核心类/函数：
 *   - showOcrBatchDialog() — 显示批量 OCR 对话框
 *   - renderBatchList() — 渲染批量列表
 *   - startBatchOcr() — 开始批量识别
 *   - exportBatchResults() — 导出批量结果
 *
 * 依赖：
 *   - app-ocr.js
 *   - app-file-batch.js
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    function bt(key, vars) {
        try {
            if (window.App && App.i18n && App.i18n.t) {
                var v = App.i18n.t(key, vars);
                if (v && v !== key) return v;
            }
        } catch (_) {}
        return key;
    }

    /**
     * 批量处理状态
     */
    const BatchState = {
        files: [],          // {path, name, status, result}
        isProcessing: false,
        isPaused: false,
        totalFiles: 0,
        completedFiles: 0,
        failedFiles: 0,
        currentFile: '',
        timerId: null
    };

    /**
     * 初始化批量 Tab
     */
    function initBatchTab() {
        const batchPanel = document.getElementById('tab-batch');
        if (!batchPanel) return;
        
        // 初始化自动跟踪按钮状态
        const btn = document.getElementById('btn-batch-autotrack');
        if (btn) {
            btn.style.opacity = _autoTrack ? '1' : '0.4';
            btn.title = bt('ocr.batch.trackTitle');
        }
        
        // 确保文件列表区域存在
        const fileList = document.getElementById('batch-file-list');
        if (!fileList) return;
        
        // 更新统计和按钮状态
        updateBatchStats();
        updateBatchActions();
        
        // 延迟填充批量解析脚本选择器（等待App.Parser加载完成）
        setTimeout(() => {
            if (typeof App !== 'undefined' && App.Parser && App.Parser.populateBatchScriptSelector) {
                App.Parser.populateBatchScriptSelector();
            }
        }, 1000);
    }

    /**
     * 引擎就绪守卫：OCR 引擎初始化/切换期间不允许发起批量任务
     * （与 ocr-panel 的 setOcrActionButtonsEnabled 通过 window.__ocrEngineReady 联动）
     */
    function ensureEngineReady() {
        if (window.__ocrEngineReady) return true;
        showBatchToast(bt('ocr.batch.engineWait'));
        return false;
    }

    /**
     * 选择文件夹
     */
    async function selectBatchFolder() {
        if (!ensureEngineReady()) return;
        try {
            const input = document.createElement('input');
            input.type = 'file';
            input.webkitdirectory = true;
            input.onchange = function(e) {
                const files = Array.from(e.target.files).filter(f =>
                    /\.(png|jpg|jpeg|bmp|tiff|tif|gif)$/i.test(f.name)
                );
                addFilesToList(files.map(f => ({
                    path: f.name,
                    name: f.name,
                    size: f.size,
                    fileObj: f,  // 保留 File 对象
                    isFromFolder: true
                })));
                if (files.length > 0) {
                    showBatchToast(`已添加 ${files.length} 个文件`);
                } else {
                    showBatchToast('文件夹中没有支持的图片文件');
                }
            };
            input.click();
        } catch (e) {
            ChatLog.error('OCR', '选择文件夹失败:', e);
            showBatchToast('选择文件夹失败');
        }
    }

    /**
     * 选择多个文件
     */
    function selectBatchFiles() {
        if (!ensureEngineReady()) return;
        const input = document.createElement('input');
        input.type = 'file';
        input.multiple = true;
        input.accept = 'image/*,.pdf,.tiff,.tif,.bmp';
        input.onchange = function(e) {
            const files = Array.from(e.target.files).map(f => ({
                path: f.name,
                name: f.name,
                size: f.size,
                fileObj: f  // 保留 File 对象用于后续读取 base64
            }));
            addFilesToList(files);
        };
        input.click();
    }

    /**
     * 扫描文件夹（通过后端）
     */
    async function scanFolder(folderPath) {
        try {
            const resultEl = document.getElementById('batch-status');
            if (resultEl) resultEl.textContent = bt('ocr.batch.scanning');

            if (window.chrome?.webview?.hostObjects?.ocrBridge) {
                const result = await window.chrome.webview.hostObjects.ocrBridge.ScanOcrFolder(folderPath, false);
                const obj = JSON.parse(result);
                if (obj.success && obj.files) {
                    const fileList = obj.files.map(fp => ({
                        path: fp,
                        name: fp.split(/[\\/]/).pop(),
                        isFromFolder: true
                    }));
                    addFilesToList(fileList);
                    if (resultEl) resultEl.textContent = bt('ocr.batch.scanDone', { n: fileList.length });
                } else {
                    if (resultEl) resultEl.textContent = bt('ocr.batch.scanFail', { msg: obj.error || bt('ocr.batch.unknownError') });
                }
            } else {
                if (resultEl) resultEl.textContent = bt('ocr.batch.noBackend');
            }
        } catch (e) {
            ChatLog.error('OCR', '扫描文件夹失败:', e);
            const resultEl = document.getElementById('batch-status');
            if (resultEl) resultEl.textContent = bt('ocr.batch.scanFail', { msg: e.message });
        }
    }

    /**
     * 添加文件到列表
     */
    function addFilesToList(files) {
        if (!files || files.length === 0) return;

        // 去重
        const existingNames = new Set(BatchState.files.map(f => f.name));
        const newFiles = files.filter(f => !existingNames.has(f.name));

        newFiles.forEach(f => {
            BatchState.files.push({
                path: f.path || f.name,
                name: f.name,
                size: f.size || 0,
                fileObj: f.fileObj || null,  // 保留 File 对象引用
                status: 'pending',  // pending, processing, completed, failed
                result: null,
                error: null
            });
        });

        BatchState.totalFiles = BatchState.files.length;
        renderFileList();
        updateBatchStats();
        updateBatchActions();

        showBatchToast(`已添加 ${newFiles.length} 个文件`);
    }

    /**
     * 渲染文件列表
     */
    function renderFileList() {
        const container = document.getElementById('batch-file-list');
        if (!container) return;

        if (BatchState.files.length === 0) {
            container.innerHTML = '<div class="batch-empty">' +
                '<div class="batch-empty-icon">&#128444;</div>' +
                '<p>' + bt('ocr.batch.emptyFiles') + '</p>' +
                '<p class="batch-empty-hint">' + bt('ocr.batchEmptyHint') + '</p>' +
                '<div style="margin-top:14px;display:flex;gap:10px;justify-content:center;">' +
                '<button class="btn-primary btn-sm" onclick="OcrBatchUI.selectBatchFiles()" style="font-size:13px;padding:8px 16px;">' + bt('ocr.batch.addFile') + '</button>' +
                '<button class="btn-secondary btn-sm" onclick="OcrBatchUI.selectBatchFolder()" style="font-size:13px;padding:8px 16px;">' + bt('ocr.batch.addFolder') + '</button>' +
                '</div></div>';
            return;
        }

        let html = '';
        // 顶部添加按钮栏
        html += '<div style="display:flex;gap:6px;padding:6px 10px;border-bottom:1px solid #f0f0f0;">';
        html += '<button class="btn-primary btn-sm" onclick="OcrBatchUI.selectBatchFiles()" style="font-size:11px;padding:3px 10px;">+ ' + bt('ocr.batch.addFile') + '</button>';
        html += '<button class="btn-secondary btn-sm" onclick="OcrBatchUI.selectBatchFolder()" style="font-size:11px;padding:3px 10px;">+ ' + bt('ocr.batch.addFolder') + '</button>';
        html += '</div>';

        BatchState.files.forEach((file, index) => {
            const statusIcon = getStatusIcon(file.status);
            const rowClass = file.status === 'failed' ? 'batch-row-failed' :
                             file.status === 'completed' ? 'batch-row-success' : '';
            const hasResult = file.result ? 'batch-has-result' : '';
            html += `<div class="batch-file-row ${rowClass} ${hasResult}" data-index="${index}" onclick="OcrBatchUI.onClickFileRow(${index})">
                <span class="batch-drag-handle" title="拖动排序"
                      onpointerdown="OcrBatchUI.startFileDrag(event, ${index})"
                      onclick="event.stopPropagation(); event.preventDefault();"
                      ondblclick="event.stopPropagation(); event.preventDefault();">⋮⋮</span>
                <span class="batch-file-icon">${statusIcon}</span>
                <span class="batch-file-name" title="${escapeHtml(file.path)}">${escapeHtml(file.name)}</span>
                <span class="batch-file-size">${file.size ? formatBatchFileSize(file.size) : ''}</span>
                <span class="batch-file-status">${getStatusText(file.status, file.error)}</span>
                <button class="batch-file-remove" onclick="event.stopPropagation();OcrBatchUI.removeBatchFile(${index})" title="${bt('ocr.batch.remove')}">&times;</button>
            </div>`;
        });

        container.innerHTML = html;
    }

    /**
     * 获取状态图标
     */
    function getStatusIcon(status) {
        const icons = {
            pending: '&#9679;',        // 灰色圆点
            processing: '&#9699;',     // 旋转
            completed: '&#10004;',     // 对勾
            failed: '&#10008;'         // 叉号
        };
        return icons[status] || icons.pending;
    }

    /**
     * 获取状态文本
     */
    function getStatusText(status, error) {
        const texts = {
            pending: bt('ocr.batch.pending'),
            processing: bt('ocr.batch.processing'),
            completed: bt('ocr.batch.completed'),
            failed: bt('ocr.batch.failedMsg', { msg: error || bt('ocr.batch.unknownError') })
        };
        return texts[status] || texts.pending;
    }

    /**
     * 更新统计信息
     */
    function updateBatchStats() {
        const totalEl = document.getElementById('batch-total');
        const completedEl = document.getElementById('batch-completed');
        const failedEl = document.getElementById('batch-failed');

        if (totalEl) totalEl.textContent = BatchState.files.length;
        if (completedEl) completedEl.textContent = BatchState.completedFiles;
        if (failedEl) failedEl.textContent = BatchState.failedFiles;
    }

    /**
     * 更新操作按钮状态
     */
    function updateBatchActions() {
        ChatLog.info('OCR', '[OCR Batch] 更新操作按钮状态...');
        const startBtn = document.getElementById('btn-batch-start');
        const pauseBtn = document.getElementById('btn-batch-pause');
        const cancelBtn = document.getElementById('btn-batch-cancel');
        const clearBtn = document.getElementById('btn-batch-clear');
        const progressContainer = document.getElementById('batch-progress-container');
        
        ChatLog.info('OCR', '[OCR Batch] 按钮元素:', {
            startBtn: !!startBtn,
            pauseBtn: !!pauseBtn,
            cancelBtn: !!cancelBtn,
            clearBtn: !!clearBtn,
            progressContainer: !!progressContainer
        });

        if (startBtn) {
            startBtn.disabled = BatchState.isProcessing || BatchState.files.length === 0;
            ChatLog.info('OCR', '[OCR Batch] 开始按钮状态:', startBtn.disabled);
        }
        if (pauseBtn) {
            pauseBtn.style.display = BatchState.isProcessing ? 'inline-flex' : 'none';
            pauseBtn.textContent = BatchState.isPaused ? bt('ocr.pdfResume') : bt('ocr.batchPause');
        }
        if (cancelBtn) {
            cancelBtn.style.display = BatchState.isProcessing ? 'inline-flex' : 'none';
        }
        if (clearBtn) clearBtn.disabled = BatchState.isProcessing;
        if (progressContainer) progressContainer.style.display = BatchState.isProcessing ? 'block' : 'none';
        
        ChatLog.info('OCR', '[OCR Batch] 操作按钮状态更新完成');
    }

    // ============================================================
    // 批量处理执行
    // ============================================================

    /** 检查文本是否为真实错误信息 */
    function isErrorResult(text) {
        if (!text || typeof text !== 'string') return true;
        return /^【错误[：：]】/.test(text.trim());
    }

    /** 根据字段配置的顺序、名称和显示开关整理结果。派生键并入对应字段。 */
    function filterVisibleFields(parsedResult, fields) {
        if (!parsedResult || typeof parsedResult !== 'object') return parsedResult;
        if (!fields || fields.length === 0) return parsedResult;

        const used = {};
        const ordered = {};
        fields.forEach(function (field) {
            if (!field || !field.key) return;
            const alias = field.key + '_from_id';
            used[field.key] = true;
            used[alias] = true;
            if (field.visible === false) return;
            let value = parsedResult[field.key];
            if (value === undefined || value === null || value === '') value = parsedResult[alias];
            if (value === undefined || value === null || value === '') return;
            ordered[field.key] = value;
        });
        Object.keys(parsedResult).forEach(function (key) {
            if (used[key]) return;
            const value = parsedResult[key];
            if (value === undefined || value === null || value === '') return;
            ordered[key] = value;
        });
        return ordered;
    }

    /**
     * 开始批量处理（支持实时预览）
     */
    async function startBatchProcessing() {
        if (BatchState.isProcessing) return;

        const pendingFiles = BatchState.files.filter(f => f.status === 'pending' || f.status === 'failed');
        if (pendingFiles.length === 0) {
            showBatchToast('没有待处理的文件');
            return;
        }

        // 保存用户设置的智能解析状态和脚本快照
        const parserCheckbox = document.getElementById('batch-parser-enabled');
        const parserScriptSelect = document.getElementById('batch-parser-script');
        const savedParserEnabled = parserCheckbox ? parserCheckbox.checked : false;
        const savedScriptId = parserScriptSelect ? parserScriptSelect.value : '';
        
        // 保存脚本快照（防止处理过程中脚本被修改或删除）
        let savedScriptSnapshot = null;
        if (savedParserEnabled && savedScriptId && typeof App !== 'undefined' && App.Parser && App.Parser.scripts) {
            const script = App.Parser.scripts.find(s => s.id === savedScriptId);
            if (script) {
                savedScriptSnapshot = {
                    id: script.id,
                    name: script.name,
                    script: script._rawScript || script.script,
                    fields: script.fields || [],
                    outputFormat: script.outputFormat || 'table'
                };
            }
        }
        
        ChatLog.info('OCR', '[Batch] 保存的智能解析设置:', { 
            enabled: savedParserEnabled, 
            script: savedScriptId,
            hasSnapshot: !!savedScriptSnapshot 
        });

        BatchState.isProcessing = true;
        BatchState.isPaused = false;
        BatchState.completedFiles = 0;
        BatchState.failedFiles = 0;
        _autoTrack = true;  // 开始处理时开启自动跟踪

        // 重置状态
        pendingFiles.forEach(f => { f.status = 'pending'; f.result = null; f.error = null; });

        // 只清空待处理文件的结果卡片，保留已完成文件的结果
        const streamContent = document.getElementById('batch-stream-content');
        if (streamContent) {
            pendingFiles.forEach(f => {
                const cardId = 'batch-stream-card-' + f.name.replace(/[^a-zA-Z0-9]/g, '_');
                const card = document.getElementById(cardId);
                if (card) card.remove();
            });
            // 重置计数器为剩余卡片数量
            _streamItemCount = streamContent.children.length;
        }

        updateBatchActions();
        renderFileList();

        const progressBar = document.getElementById('batch-progress-bar');
        const progressText = document.getElementById('batch-progress-text');
        const statusEl = document.getElementById('batch-status');

        if (progressBar) progressBar.style.width = '0%';
        if (progressText) progressText.textContent = '0%';
        if (statusEl) statusEl.textContent = bt('ocr.batch.running', { n: 0, total: pendingFiles.length });

        // 同步更新固定状态栏（强制执行）
        try {
            if (typeof App !== 'undefined' && App.fixedStatusBar) {
                if (!App.fixedStatusBar.isInitialized) {
                    App.fixedStatusBar.init();
                }
                App.fixedStatusBar.setLoading(`批量OCR: 0/${pendingFiles.length}`, 0);
            }
        } catch(e) {
            ChatLog.warn('OCR', '[FixedStatusBar] 更新失败:', e);
        }

        const total = pendingFiles.length;

        for (let i = 0; i < pendingFiles.length; i++) {
            // 暂停检查
            while (BatchState.isPaused) {
                await sleep(200);
            }

            if (!BatchState.isProcessing) break; // 已取消

            const file = pendingFiles[i];
            const fileIndex = BatchState.files.indexOf(file); // 获取在BatchState.files中的真实索引
            file.status = 'processing';
            BatchState.currentFile = file.name;
            // 为当前文件保存脚本快照
            file.scriptSnapshot = savedScriptSnapshot;

            // 更新 UI
            renderFileList();
            if (statusEl) statusEl.textContent = bt('ocr.batch.fileProgress', { n: i + 1, total: total, name: file.name });

            // 同步更新固定状态栏 - 进度（强制执行）
            try {
                if (typeof App !== 'undefined' && App.fixedStatusBar) {
                    const percent = Math.round(((i + 1) / total) * 100);
                    App.fixedStatusBar.updateProgress(percent, `正在识别: ${file.name}`);
                }
            } catch(e) {
                ChatLog.warn('OCR', '[FixedStatusBar] 更新失败:', e);
            }

            // *** 流式结果：处理开始时追加一条加载中的条目 ***
            appendStreamResult(fileIndex, null);

            try {
                const result = await recognizeSingleFile(file);
                ChatLog.info('OCR', '[Batch] OCR结果:', file.name, result ? result.substring(0, 100) + '...' : 'empty');
                
                // 检查结果是否为错误文本
                if (isErrorResult(result)) {
                    throw new Error(result.replace(/^【错误[：:】]/, '').trim());
                }
                file.status = 'completed';
                file.result = result;
                BatchState.completedFiles++;
                ChatLog.info('OCR', '[Batch] 调用 updateStreamResult:', fileIndex, 'completed');
                // *** 流式结果：完成后更新该条目为完成状态（使用快照解析） ***
                updateStreamResult(fileIndex, 'completed', result, null, savedParserEnabled, savedScriptId);
            } catch (e) {
                file.status = 'failed';
                file.error = e.message || '处理失败';
                BatchState.failedFiles++;
                // *** 流式结果：更新该条目为失败状态 ***
                updateStreamResult(fileIndex, 'failed', null, e.message);
            }

            // 更新进度
            const percent = Math.round(((i + 1) / total) * 100);
            if (progressBar) progressBar.style.width = percent + '%';
            if (progressText) progressText.textContent = percent + '%';

            updateBatchStats();
            renderFileList();
        }

        BatchState.isProcessing = false;
        BatchState.currentFile = '';

        if (statusEl) {
            statusEl.textContent = bt('ocr.batch.doneSummary', { ok: BatchState.completedFiles, bad: BatchState.failedFiles });
        }

        // 同步更新固定状态栏 - 完成状态（强制执行）
        try {
            if (typeof App !== 'undefined' && App.fixedStatusBar) {
                if (BatchState.failedFiles > 0) {
                    App.fixedStatusBar.setSuccess(`OCR完成: ${BatchState.completedFiles} 成功, ${BatchState.failedFiles} 失败`, 3000);
                } else {
                    App.fixedStatusBar.setSuccess(`全部 ${BatchState.completedFiles} 个文件识别完成`, 3000);
                }
            }
        } catch(e) {
            ChatLog.warn('OCR', '[FixedStatusBar] 更新失败:', e);
        }

        updateBatchActions();

        showBatchToast('批量处理完成');
    }

    /**
     * 解析单个文件（带指定脚本）
     */
    function parseSingleFileWithScript(index, scriptId) {
        if (!scriptId) return; // 用户选择"快速解析..."时不执行
        
        const file = BatchState.files[index];
        if (!file || file.status !== 'completed' || !file.result) {
            showBatchToast('文件未识别完成');
            return;
        }

        // 从App.Parser.scripts中查找脚本
        const scripts = (typeof App !== 'undefined' && App.Parser) ? App.Parser.scripts || [] : [];
        const script = scripts.find(s => s.id === scriptId);

        if (!script || !(script._rawScript || script.script)) {
            showBatchToast('未找到解析脚本');
            return;
        }

        try {
            const scriptCode = script._rawScript || script.script;
            const fields = script.fields || [];
            const format = script.outputFormat || 'table';

            // 策略2: 执行整个脚本，返回parse函数
            const wrappedAsDeclaration = `
                "use strict";
                const FIELDS = ${JSON.stringify(fields)};
                ${scriptCode};
                if (typeof parse === 'function') return parse;
            `;

            const func = new Function(wrappedAsDeclaration);
            const parseFunc = func();

            if (typeof parseFunc !== 'function') {
                showBatchToast('脚本没有返回parse函数');
                return;
            }

            const parsedResult = parseFunc(file.result);
            file.parsedResult = parsedResult;

            // 根据字段 visible 配置过滤，用于渲染和后续导出
            const displayResult = filterVisibleFields(parsedResult, fields);

            // 直接更新DOM显示解析结果
            const cardId = 'batch-stream-card-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
            const card = document.getElementById(cardId);
            if (card) {
                const body = card.querySelector('.stream-card-body');
                if (body) {
                    const renderId = 'batch-stream-render-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
                    body.innerHTML = '<div id="' + renderId + '"></div>';

                    // 根据格式渲染（使用过滤后的可见字段）
                    if (format === 'table') {
                        let html = '<table class="parse-result-table"><thead><tr><th>字段名</th><th>值</th></tr></thead><tbody>';
                        Object.entries(displayResult).forEach(([key, value]) => {
                            const fieldName = (fields.find(f => f.key === key) || {}).name || key;
                            html += `<tr><td><strong>${escapeHtml(fieldName)}</strong></td><td>${escapeHtml(String(value))}</td></tr>`;
                        });
                        html += '</tbody></table>';
                        document.getElementById(renderId).innerHTML = html;
                    } else if (format === 'key-value') {
                        let html = '<div class="parse-kv-list">';
                        Object.entries(displayResult).forEach(([key, value]) => {
                            const fieldName = (fields.find(f => f.key === key) || {}).name || key;
                            html += `<div class="parse-kv-item"><span class="parse-kv-key">${escapeHtml(fieldName)}</span><span class="parse-kv-value">${escapeHtml(String(value))}</span></div>`;
                        });
                        html += '</div>';
                        document.getElementById(renderId).innerHTML = html;
                    } else {
                        document.getElementById(renderId).innerHTML = '<pre style="margin:0;padding:12px;background:white;border-radius:4px;font-size:12px;line-height:1.5;overflow:auto;">' + escapeHtml(JSON.stringify(displayResult, null, 2)) + '</pre>';
                    }
                }
            }

            showBatchToast('解析成功: ' + file.name);
        } catch (e) {
            ChatLog.error('OCR', '[Batch] 单文件解析失败:', file.name, e);
            showBatchToast('解析失败: ' + e.message);
        }
    }

    /**
     * 解析单个文件（使用全局设置）
     */
    function parseSingleFile(index) {
        const file = BatchState.files[index];
        if (!file || file.status !== 'completed' || !file.result) {
            showBatchToast('文件未识别完成');
            return;
        }

        const parserCheckbox = document.getElementById('batch-parser-enabled');
        const parserScriptSelect = document.getElementById('batch-parser-script');
        const parserEnabled = parserCheckbox ? parserCheckbox.checked : false;
        const scriptId = parserScriptSelect ? parserScriptSelect.value : '';

        if (!parserEnabled || !scriptId) {
            showBatchToast('请启用智能解析并选择脚本');
            return;
        }

        // 从App.Parser.scripts中查找脚本
        const scripts = (typeof App !== 'undefined' && App.Parser) ? App.Parser.scripts || [] : [];
        const script = scripts.find(s => s.id === scriptId);

        if (!script || !(script._rawScript || script.script)) {
            showBatchToast('未找到解析脚本');
            return;
        }

        try {
            const scriptCode = script._rawScript || script.script;
            const fields = script.fields || [];
            const format = script.outputFormat || 'table';

            // 策略2: 执行整个脚本，返回parse函数
            const wrappedAsDeclaration = `
                "use strict";
                const FIELDS = ${JSON.stringify(fields)};
                ${scriptCode};
                if (typeof parse === 'function') return parse;
            `;

            const func = new Function(wrappedAsDeclaration);
            const parseFunc = func();

            if (typeof parseFunc !== 'function') {
                showBatchToast('脚本没有返回parse函数');
                return;
            }

            const parsedResult = parseFunc(file.result);
            file.parsedResult = parsedResult;

            // 根据字段 visible 配置过滤，用于渲染和后续导出
            const displayResult = filterVisibleFields(parsedResult, fields);

            // 直接更新DOM显示解析结果
            const cardId = 'batch-stream-card-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
            const card = document.getElementById(cardId);
            if (card) {
                const body = card.querySelector('.stream-card-body');
                if (body) {
                    const renderId = 'batch-stream-render-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
                    body.innerHTML = '<div id="' + renderId + '"></div>';

                    // 根据格式渲染（使用过滤后的可见字段）
                    if (format === 'table') {
                        let html = '<table class="parse-result-table"><thead><tr><th>字段名</th><th>值</th></tr></thead><tbody>';
                        Object.entries(displayResult).forEach(([key, value]) => {
                            const fieldName = (fields.find(f => f.key === key) || {}).name || key;
                            html += `<tr><td><strong>${escapeHtml(fieldName)}</strong></td><td>${escapeHtml(String(value))}</td></tr>`;
                        });
                        html += '</tbody></table>';
                        document.getElementById(renderId).innerHTML = html;
                    } else if (format === 'key-value') {
                        let html = '<div class="parse-kv-list">';
                        Object.entries(displayResult).forEach(([key, value]) => {
                            const fieldName = (fields.find(f => f.key === key) || {}).name || key;
                            html += `<div class="parse-kv-item"><span class="parse-kv-key">${escapeHtml(fieldName)}</span><span class="parse-kv-value">${escapeHtml(String(value))}</span></div>`;
                        });
                        html += '</div>';
                        document.getElementById(renderId).innerHTML = html;
                    } else {
                        document.getElementById(renderId).innerHTML = '<pre style="margin:0;padding:12px;background:white;border-radius:4px;font-size:12px;line-height:1.5;overflow:auto;">' + escapeHtml(JSON.stringify(displayResult, null, 2)) + '</pre>';
                    }
                }
            }

            showBatchToast('解析成功: ' + file.name);
        } catch (e) {
            ChatLog.error('OCR', '[Batch] 单文件解析失败:', file.name, e);
            showBatchToast('解析失败: ' + e.message);
        }
    }

    /**
     * 批量解析所有已完成的文件
     */
    async function parseCompletedFiles() {
        const completedFiles = BatchState.files.filter(f => f.status === 'completed' && f.result);
        if (completedFiles.length === 0) {
            showBatchToast(bt('ocr.batch.noFiles'));
            return;
        }

        const statusEl = document.getElementById('batch-status');
        const scripts = (typeof App !== 'undefined' && App.Parser) ? App.Parser.scripts || [] : [];
        
        let successCount = 0;
        let failCount = 0;

        for (let i = 0; i < completedFiles.length; i++) {
            const file = completedFiles[i];
            const fileIndex = BatchState.files.indexOf(file);
            
            // 从当前卡片的脚本选择器获取脚本ID
            const cardId = 'batch-stream-card-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
            const card = document.getElementById(cardId);
            if (!card) continue;
            
            const select = card.querySelector('select');
            const scriptId = select ? select.value : '';
            
            if (!scriptId) {
                ChatLog.info('OCR', '[Batch] 跳过未选择脚本的文件:', file.name);
                continue;
            }

            if (statusEl) statusEl.textContent = bt('ocr.batch.parsing', { n: i + 1, total: completedFiles.length, name: file.name });

            try {
                // 从App.Parser.scripts中查找脚本
                const script = scripts.find(s => s.id === scriptId);
                if (!script || !(script._rawScript || script.script)) {
                    ChatLog.error('OCR', '[Batch] 未找到脚本:', scriptId);
                    failCount++;
                    continue;
                }

                const scriptCode = script._rawScript || script.script;
                const fields = script.fields || [];
                const format = script.outputFormat || 'table';

                // 策略2: 执行整个脚本，返回parse函数
                const wrappedAsDeclaration = `
                    "use strict";
                    const FIELDS = ${JSON.stringify(fields)};
                    ${scriptCode};
                    if (typeof parse === 'function') return parse;
                `;

                const func = new Function(wrappedAsDeclaration);
                const parseFunc = func();

                if (typeof parseFunc !== 'function') {
                    ChatLog.error('OCR', '[Batch] 脚本没有返回parse函数:', script.name);
                    failCount++;
                    continue;
                }

                const parsedResult = parseFunc(file.result);
                file.parsedResult = parsedResult;

                // 根据字段 visible 配置过滤，用于渲染和后续导出
                const displayResult = filterVisibleFields(parsedResult, fields);

                // 直接更新DOM显示解析结果
                const body = card.querySelector('.stream-card-body');
                if (body) {
                    const renderId = 'batch-stream-render-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
                    body.innerHTML = '<div id="' + renderId + '"></div>';

                    // 根据格式渲染（使用过滤后的可见字段）
                    if (format === 'table') {
                        let html = '<table class="parse-result-table"><thead><tr><th>字段名</th><th>值</th></tr></thead><tbody>';
                        Object.entries(displayResult).forEach(([key, value]) => {
                            const fieldName = (fields.find(f => f.key === key) || {}).name || key;
                            html += `<tr><td><strong>${escapeHtml(fieldName)}</strong></td><td>${escapeHtml(String(value))}</td></tr>`;
                        });
                        html += '</tbody></table>';
                        document.getElementById(renderId).innerHTML = html;
                    } else if (format === 'key-value') {
                        let html = '<div class="parse-kv-list">';
                        Object.entries(displayResult).forEach(([key, value]) => {
                            const fieldName = (fields.find(f => f.key === key) || {}).name || key;
                            html += `<div class="parse-kv-item"><span class="parse-kv-key">${escapeHtml(fieldName)}</span><span class="parse-kv-value">${escapeHtml(String(value))}</span></div>`;
                        });
                        html += '</div>';
                        document.getElementById(renderId).innerHTML = html;
                    } else {
                        document.getElementById(renderId).innerHTML = '<pre style="margin:0;padding:12px;background:white;border-radius:4px;font-size:12px;line-height:1.5;overflow:auto;">' + escapeHtml(JSON.stringify(displayResult, null, 2)) + '</pre>';
                    }
                }
                successCount++;
            } catch (parseError) {
                ChatLog.error('OCR', '[Batch] 智能解析失败:', file.name, parseError);
                failCount++;
            }
        }

        if (statusEl) statusEl.textContent = bt('ocr.batch.doneSummary', { ok: BatchState.completedFiles, bad: BatchState.failedFiles });
        showBatchToast(bt('ocr.batch.parseDone', { ok: successCount, bad: failCount }));
    }

    // ============================================================
    // 结果流（所有结果累计显示）
    // ============================================================

    let _selectedIndex = -1;
    let _autoTrack = true;
    let _streamItemCount = 0;

    /** 将 File 对象读取为 base64 数据 URL */
    function readFileAsBase64(fileObj) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => reject(new Error('读取文件失败'));
            reader.readAsDataURL(fileObj);
        });
    }

    /**
     * 读取批量识别使用的输出格式（与 OCR 面板共享 localStorage 键 ocr_format_setting）
     * 旧版批量固定 plaintext，忽略用户在面板选择的格式
     */
    function getBatchOutputFormat() {
        const validFormats = ['plaintext', 'markdown', 'html', 'excel'];
        try {
            const saved = localStorage.getItem('ocr_format_setting');
            if (saved && validFormats.includes(saved)) return saved;
        } catch (e) {
            // localStorage 不可用（隐私模式等）时回退默认格式
        }
        return 'plaintext';
    }

    /**
     * 识别单个文件
     * 使用与单文件 OCR 相同的流程：SaveBase64Image → StartOcrRecognitionWithFormat → Poll
     * 输出格式跟随面板的格式选择（localStorage: ocr_format_setting，默认 plaintext）
     */
    async function callBatchOcr(bridgeName, hostName, args) {
        // 本地 OCR 每日免费额度（云通道 / 已激活不计数）。详情见 js/ocr-quota.js
        var quota = window.SmartPilotOcrQuota;
        if (quota && quota.gate({ method: bridgeName, engine: window.__currentOcrEngine || '' }) === false) {
            throw new Error('OCR_QUOTA_EXCEEDED');
        }
        try {
            const bridge = window.chrome?.webview?.hostObjects?.excelBridge;
            if (bridge) return await bridge[bridgeName](...args);
            if (typeof window.ocrHostCall === 'function') return await window.ocrHostCall(hostName, args);
            throw new Error('本机识图服务未启动');
        } catch (e) {
            // 调用失败：退还这次额度（只为真正启动的识别计费）
            if (quota) quota.refund({ method: bridgeName, engine: window.__currentOcrEngine || '' });
            throw e;
        }
    }

    async function recognizeSingleFile(file) {
        let filePath = file.path;

        // 如果有 File 对象（来自 input[type=file]），先保存为临时文件
        if (file.fileObj) {
            const base64 = await readFileAsBase64(file.fileObj);
            const saveResult = await callBatchOcr(
                'SaveBase64Image',
                'OcrServiceManager.SaveBase64Image',
                [base64, file.name]
            );
            const saveObj = JSON.parse(saveResult);
            if (!saveObj.success) throw new Error('无法保存图片文件: ' + (saveObj.error || ''));
            filePath = saveObj.path;
        }

        // 输出格式跟随面板选择（默认 plaintext，与旧行为一致）
        const outputFormat = getBatchOutputFormat();
        const recognitionMode = 'auto';

        // 启动 OCR 识别（带格式）。Excel 里走加载项桥，否则走本机识图服务。
        const startResult = await callBatchOcr(
            'StartOcrRecognitionWithFormat',
            'OcrBridge.StartOcrRecognitionWithFormat',
            [filePath, recognitionMode, outputFormat]
        );
        const startObj = JSON.parse(startResult);
        if (!startObj.started) {
            // 启动失败退还额度（与工作台 recognizeImage 同口径：只为真正启动的识别计费）
            if (window.SmartPilotOcrQuota) {
                window.SmartPilotOcrQuota.refund({ method: 'StartOcrRecognitionWithFormat', engine: window.__currentOcrEngine || '' });
            }
            throw new Error(startObj.error || 'OCR 启动失败');
        }

        // 轮询结果
        const maxWait = 600 * 1000;
        const pollStart = Date.now();

        while (true) {
            await sleep(500);
            if (Date.now() - pollStart > maxWait) throw new Error('识别超时');

            const pollResult = await callBatchOcr(
                'GetOcrResultWithFormat',
                'OcrBridge.GetOcrResultWithFormat',
                [outputFormat]
            );
            const pollObj = JSON.parse(pollResult);
            if (pollObj.ready) {
                if (pollObj.success) return pollObj.text;
                throw new Error(pollObj.error || '识别失败');
            }
        }
    }

    /**
     * 暂停批量处理
     */
    function pauseBatch() {
        BatchState.isPaused = !BatchState.isPaused;
        updateBatchActions();
        const statusEl = document.getElementById('batch-status');
        if (statusEl) {
            statusEl.textContent = BatchState.isPaused ? bt('ocr.batch.paused') : bt('ocr.batch.resuming');
        }
    }

    /**
     * 取消批量处理
     */
    function cancelBatch() {
        BatchState.isProcessing = false;
        BatchState.isPaused = false;
        updateBatchActions();
        renderFileList();

        const statusEl = document.getElementById('batch-status');
        if (statusEl) statusEl.textContent = bt('ocr.batch.cancelled');
        showBatchToast(bt('ocr.batch.cancelledToast'));
    }

    /**
     * 移除单个文件
     */
    function removeBatchFile(index) {
        if (BatchState.isProcessing) return;
        if (index >= 0 && index < BatchState.files.length) {
            const file = BatchState.files[index];
            
            // 只清除被删除文件对应的结果卡片
            const cardId = 'batch-stream-card-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
            const card = document.getElementById(cardId);
            if (card) {
                card.remove();
                _streamItemCount--;
            }
            
            // 从文件列表中移除
            BatchState.files.splice(index, 1);
            BatchState.totalFiles = BatchState.files.length;
            
            // 重新统计
            BatchState.completedFiles = BatchState.files.filter(f => f.status === 'completed').length;
            BatchState.failedFiles = BatchState.files.filter(f => f.status === 'failed').length;
            
            renderFileList();
            updateBatchStats();
            updateBatchActions();
        }
    }

    /**
     * 清空文件列表
     */
    function clearBatchList() {
        if (BatchState.isProcessing) return;
        if (BatchState.files.length === 0) return;
        if (!confirm('确定要清空所有文件吗？')) return;

        clearStreamResults();
        BatchState.files = [];
        BatchState.totalFiles = 0;
        BatchState.completedFiles = 0;
        BatchState.failedFiles = 0;
        renderFileList();
        updateBatchStats();
        updateBatchActions();

        const progressBar = document.getElementById('batch-progress-bar');
        const progressText = document.getElementById('batch-progress-text');
        if (progressBar) progressBar.style.width = '0%';
        if (progressText) progressText.textContent = '0%';
    }

    /**
     * 批量导出结果到 Excel（智能插入：检测现有表格并追加数据）
     */
    async function exportBatchToExcel() {
        // 优先使用有解析结果的文件
        const parsedFiles = BatchState.files.filter(f =>
            f.status === 'completed' && f.parsedResult && Object.keys(f.parsedResult).length > 0
        );
        
        // 如果没有解析结果，使用OCR原始文本
        const ocrFiles = BatchState.files.filter(f =>
            f.status === 'completed' && f.result && !isErrorResult(f.result) && !f.parsedResult
        );
        
        const failed = BatchState.files.filter(f =>
            f.status === 'failed' || (f.status === 'completed' && isErrorResult(f.result))
        );

        if (parsedFiles.length === 0 && ocrFiles.length === 0) {
            showBatchToast('没有可导出的成功结果');
            return;
        }

        try {
            let rows = [];
            let useParsedData = parsedFiles.length > 0;
            
            if (useParsedData) {
                // 使用解析后的结构化数据（横向表格）
                rows = parsedFiles.map(f => {
                    const row = {};
                    
                    // 获取当前文件的解析脚本
                    const cardId = 'batch-stream-card-' + f.name.replace(/[^a-zA-Z0-9]/g, '_');
                    const card = document.getElementById(cardId);
                    let scriptId = '';
                    if (card) {
                        const select = card.querySelector('select');
                        scriptId = select ? select.value : '';
                    }
                    
                    // 查找脚本定义
                    let fields = [];
                    if (scriptId && typeof App !== 'undefined' && App.Parser && App.Parser.scripts) {
                        const script = App.Parser.scripts.find(s => s.id === scriptId);
                        if (script && script.fields) {
                            fields = script.fields;
                        }
                    }
                    
                    // 根据字段 visible 配置过滤，再使用中文显示名作为列名
                    const displayResult = filterVisibleFields(f.parsedResult, fields);
                    Object.entries(displayResult).forEach(([key, value]) => {
                        // 查找字段的中文显示名
                        const field = fields.find(f => f.key === key);
                        const displayName = field ? field.name : key;
                        row[displayName] = value || '';
                    });
                    
                    // 最后添加文件名和状态
                    row['文件名'] = f.name;
                    row['状态'] = '成功';
                    return row;
                });
            } else {
                // 使用OCR原始文本
                rows = ocrFiles.map(f => ({
                    '文件名': f.name,
                    '识别文本': f.result || '',
                    '状态': '成功'
                }));
            }
            
            // 添加失败记录
            failed.forEach(f => {
                rows.push({
                    '文件名': f.name,
                    '识别文本': f.result || '',
                    '状态': f.error || '失败'
                });
            });

            const bridge = batchExcelBridge();
            if (!bridge) {
                downloadBatchRows(rows);
                return;
            }

            // 调用新的智能插入方法
            const result = await bridge.InsertParsedData(JSON.stringify(rows));
            const resultObj = JSON.parse(result);
            
            if (!resultObj.success) {
                throw new Error(resultObj.error || resultObj.message || '插入失败');
            }
            
            // 批量导出成功弹窗已移除
        } catch (e) {
            showBatchToast('导出失败: ' + e.message);
        }
    }

    function batchExcelBridge() {
        if (window.BridgeAccessor && BridgeAccessor._usingLocalBridge) return null;
        try {
            const host = window.chrome && window.chrome.webview && window.chrome.webview.hostObjects
                && window.chrome.webview.hostObjects.excelBridge;
            if (host && typeof host.InsertParsedData === 'function') return host;
        } catch (e) { }
        try {
            if (window.BridgeAccessor && typeof BridgeAccessor.getBridge === 'function') {
                const wrapped = BridgeAccessor.getBridge();
                if (wrapped && typeof wrapped.InsertParsedData === 'function' && !BridgeAccessor._usingLocalBridge) {
                    return wrapped;
                }
            }
        } catch (e) { }
        return null;
    }

    function downloadBatchRows(rows) {
        const bytes = (window.XlsxTable && XlsxTable.build)
            ? XlsxTable.build(rows)
            : new Uint8Array(0);
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        const stamp = new Date();
        const pad = function (n) { return (n < 10 ? '0' : '') + n; };
        link.href = url;
        link.download = 'ocr_batch_' + stamp.getFullYear() + pad(stamp.getMonth() + 1) + pad(stamp.getDate()) + '_' + pad(stamp.getHours()) + pad(stamp.getMinutes()) + pad(stamp.getSeconds()) + '.xlsx';
        link.click();
        URL.revokeObjectURL(url);
        showBatchToast('当前页面不在 Excel 里，已下载表格文件');
    }

    function toggleAutoTrack() {
        _autoTrack = !_autoTrack;
        const btn = document.getElementById('btn-batch-autotrack');
        if (btn) {
            btn.style.opacity = _autoTrack ? '1' : '0.4';
            btn.title = _autoTrack ? '自动滚动到最新结果' : '已暂停自动滚动';
        }
        showBatchToast(_autoTrack ? '已开启自动跟踪' : '已暂停自动跟踪');
    }

    var _batchDrag = null;

    function remapBatchIndex(index, from, to) {
        if (index == null || index < 0) return index;
        if (index === from) return to;
        if (from < to && index > from && index <= to) return index - 1;
        if (to < from && index >= to && index < from) return index + 1;
        return index;
    }

    function batchDropIndex(clientY) {
        const container = document.getElementById('batch-file-list');
        const rows = container ? container.querySelectorAll('.batch-file-row') : [];
        for (let i = 0; i < rows.length; i++) {
            const rect = rows[i].getBoundingClientRect();
            if (clientY < rect.top + rect.height / 2) return parseInt(rows[i].dataset.index, 10);
        }
        return rows.length;
    }

    function clearBatchDragMarkers() {
        document.querySelectorAll('#batch-file-list .batch-file-row').forEach(function (el) {
            el.classList.remove('dragging', 'drag-over-before', 'drag-over-after');
        });
    }

    function reorderStreamCards() {
        const container = document.getElementById('batch-stream-content');
        if (!container) return;
        BatchState.files.forEach(function (file, index) {
            const cardId = 'batch-stream-card-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
            const card = document.getElementById(cardId);
            if (!card) return;
            card.dataset.fileIndex = String(index);
            container.appendChild(card);
        });
    }

    function onBatchDragMove(event) {
        const drag = _batchDrag;
        if (!drag || event.pointerId !== drag.pointerId) return;
        if (!drag.moved && Math.abs(event.clientY - drag.startY) < 4) return;
        drag.moved = true;
        if (drag.row) drag.row.classList.add('dragging');
        const scroller = document.getElementById('batch-file-list');
        if (scroller) {
            const rect = scroller.getBoundingClientRect();
            if (event.clientY < rect.top + 24) scroller.scrollTop -= 12;
            else if (event.clientY > rect.bottom - 24) scroller.scrollTop += 12;
        }
        const rows = scroller ? scroller.querySelectorAll('.batch-file-row') : [];
        rows.forEach(function (el) {
            el.classList.remove('drag-over-before', 'drag-over-after');
        });
        let marked = false;
        for (let i = 0; i < rows.length; i++) {
            const rect = rows[i].getBoundingClientRect();
            if (event.clientY < rect.top + rect.height / 2) {
                rows[i].classList.add('drag-over-before');
                marked = true;
                break;
            }
        }
        if (!marked && rows.length) rows[rows.length - 1].classList.add('drag-over-after');
    }

    function onBatchDragEnd(event) {
        const drag = _batchDrag;
        if (!drag) return;
        if (event && drag.pointerId != null && event.pointerId !== drag.pointerId) return;
        const handle = drag.handle;
        if (handle) {
            handle.removeEventListener('pointermove', onBatchDragMove);
            handle.removeEventListener('pointerup', onBatchDragEnd);
            handle.removeEventListener('pointercancel', onBatchDragEnd);
        }
        _batchDrag = null;
        clearBatchDragMarkers();
        if (!drag.moved) return;
        const files = BatchState.files;
        const from = drag.from;
        let to = batchDropIndex(event.clientY);
        if (to > from) to -= 1;
        if (to < 0) to = 0;
        if (to > files.length - 1) to = files.length - 1;
        if (to === from || from < 0 || from >= files.length) return;
        const moved = files.splice(from, 1)[0];
        files.splice(to, 0, moved);
        _selectedIndex = remapBatchIndex(_selectedIndex, from, to);
        reorderStreamCards();
        renderFileList();
        const row = document.querySelector('.batch-file-row[data-index="' + _selectedIndex + '"]');
        if (row) row.classList.add('batch-row-selected');
    }

    function startFileDrag(event, index) {
        if (event.button != null && event.button !== 0) return;
        if (BatchState.isProcessing) return;
        event.preventDefault();
        event.stopPropagation();
        const files = BatchState.files;
        if (!files || index < 0 || index >= files.length || files.length < 2) return;
        const handle = event.currentTarget;
        const row = handle && handle.closest ? handle.closest('.batch-file-row') : null;
        try { handle.setPointerCapture(event.pointerId); } catch (e) { }
        _batchDrag = {
            from: index,
            startY: event.clientY,
            moved: false,
            row: row,
            handle: handle,
            pointerId: event.pointerId
        };
        handle.addEventListener('pointermove', onBatchDragMove);
        handle.addEventListener('pointerup', onBatchDragEnd);
        handle.addEventListener('pointercancel', onBatchDragEnd);
    }

    function onClickFileRow(index) {
        const file = BatchState.files[index];
        if (!file) return;
        
        const cardId = 'batch-stream-card-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
        const card = document.getElementById(cardId);
        if (card) {
            card.scrollIntoView({ behavior: 'smooth', block: 'center' });
            card.style.transition = 'background 0.3s';
            card.style.background = '#e3f2fd';
            setTimeout(() => { card.style.background = ''; }, 1500);
        }
        document.querySelectorAll('.batch-file-row').forEach(el => el.classList.remove('batch-row-selected'));
        const row = document.querySelector(`.batch-file-row[data-index="${index}"]`);
        if (row) row.classList.add('batch-row-selected');
        _selectedIndex = index;
    }

    function appendStreamResult(index) {
        const container = document.getElementById('batch-stream-content');
        const stream = document.getElementById('batch-result-stream');
        if (!container || !stream) return;

        const file = BatchState.files[index];
        if (!file) return;

        stream.style.display = 'flex';

        // 使用文件名作为唯一标识，而不是索引
        const cardId = 'batch-stream-card-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
        if (document.getElementById(cardId)) return;

        const card = document.createElement('div');
        card.id = cardId;
        card.dataset.fileIndex = index; // 保存索引引用
        card.className = 'stream-card stream-card-processing';
        card.style.cssText = 'border:1px solid #f0f0f0;border-radius:6px;margin-bottom:6px;overflow:hidden;';

        card.innerHTML = '<div class="stream-card-header" style="display:flex;align-items:center;gap:8px;padding:6px 10px;background:#fafbfc;border-bottom:1px solid #f0f0f0;">' +
            '<span class="stream-spinner" style="display:inline-block;width:12px;height:12px;border:2px solid #e0e0e0;border-top-color:#4a90d9;border-radius:50%;animation:batch-spin 0.8s linear infinite;"></span>' +
            '<span style="font-size:12px;font-weight:600;color:#333;flex:1;">' + escapeHtml(file.name) + '</span>' +
            '<span style="font-size:11px;color:#999;"> 识别中</span>' +
        '</div>' +
        '<div class="stream-card-body" style="padding:8px 10px;font-size:11px;color:#999;">加载图片中</div>';

        container.appendChild(card);
        _streamItemCount++;
        if (_autoTrack) { container.scrollTop = container.scrollHeight; }
    }

    function updateStreamResult(index, status, resultText, errorMsg, parserEnabled = false, scriptId = '') {
        const file = BatchState.files[index];
        if (!file) return;
        
        // 使用文件名查找卡片
        const cardId = 'batch-stream-card-' + file.name.replace(/[^a-zA-Z0-9]/g, '_');
        const card = document.getElementById(cardId);
        if (!card) return;

        const header = card.querySelector('.stream-card-header');
        const body = card.querySelector('.stream-card-body');
        if (!header || !body) return;

        if (status === 'completed' && resultText && !isErrorResult(resultText)) {
            card.style.borderColor = '#c8e6c9';
            
            // 构建脚本选项HTML
            let scriptOptions = '<option value="">快速解析...</option>';
            if (typeof App !== 'undefined' && App.Parser && App.Parser.scripts) {
                App.Parser.scripts.forEach(s => {
                    scriptOptions += `<option value="${s.id}">${escapeHtml(s.name)}</option>`;
                });
            }
            
            header.innerHTML = '<span style="font-size:12px;font-weight:600;color:#2e7d32;flex:1;">' +
                escapeHtml(file.name) + '</span>' +
                '<select onchange="OcrBatchUI.parseSingleFileWithScript(" + index + ", this.value)" style="padding:2px 6px;font-size:11px;background:white;border:1px solid #e0e0e0;border-radius:3px;cursor:pointer;margin-right:6px;max-width:120px;">' +
                scriptOptions +
                '</select>' +
                '<span style="font-size:11px;color:#4caf50;">✅ 成功</span>';

            ChatLog.info('OCR', '[Batch] 智能解析状态:', {
                enabled: parserEnabled,
                scriptId: scriptId
            });
            
            if (parserEnabled && typeof App !== 'undefined' && App.Parser) {
                // 使用智能解析脚本渲染
                if (scriptId) {
                    // 优先使用文件对象的脚本快照，其次从App.Parser.scripts中查找
                    let script = null;
                    if (file.scriptSnapshot) {
                        // 使用快照（防止处理过程中脚本被修改）
                        script = file.scriptSnapshot;
                        ChatLog.info('OCR', '[Batch] 使用脚本快照:', script.name);
                    } else {
                        // 从App.Parser.scripts中查找
                        const scripts = App.Parser.scripts || [];
                        script = scripts.find(s => s.id === scriptId);
                        ChatLog.info('OCR', '[Batch] 从App.Parser查找脚本:', scriptId);
                    }
                    
                    if (script && (script.script || script._rawScript)) {
                        // 直接使用App.Parser的内部函数来创建解析函数
                        try {
                            const scriptCode = script._rawScript || script.script;
                            const fields = script.fields || [];
                            
                            ChatLog.info('OCR', '[Batch] 使用App.Parser逻辑创建解析函数');
                            
                            // 策略2: 处理函数声明格式 (function parse(text) {...})
                            const wrappedAsDeclaration = `
                                "use strict";
                                const FIELDS = ${JSON.stringify(fields)};
                                ${scriptCode};
                                if (typeof parse === 'function') return parse;
                            `;
                            
                            const func = new Function(wrappedAsDeclaration);
                            const parseFunc = func();
                            
                            if (typeof parseFunc !== 'function') {
                                throw new Error('脚本没有返回parse函数');
                            }
                            
                            ChatLog.info('OCR', '[Batch] 解析函数创建成功');
                            
                            if (parseFunc) {
                                // 预处理OCR文本
                                let cleanText = resultText;
                                if (/\[\d+\]\s*text=/.test(resultText)) {
                                    const textMatches = resultText.match(/\[\d+\]\s*text="([^"]*)"/g);
                                    if (textMatches) {
                                        cleanText = textMatches.map(function(m) {
                                            var match = m.match(/text="([^"]*)"/);
                                            return match ? match[1] : '';
                                        }).filter(function(t) { return t; }).join('\n');
                                    }
                                }
                                
                                const FIELDS = fields.map(f => f.key || f);
                                const parseResult = parseFunc(cleanText, FIELDS);
                                const format = script.outputFormat || 'table';
                                const renderId = 'batch-stream-render-' + index;
                                body.innerHTML = '<div id="' + renderId + '"></div>';
                                
                                // 根据格式渲染
                                if (format === 'table') {
                                    // 渲染为表格
                                    let html = '<table class="parse-result-table"><thead><tr><th>字段名</th><th>值</th></tr></thead><tbody>';
                                    Object.entries(parseResult).forEach(([key, value]) => {
                                        const fieldName = (fields.find(f => f.key === key) || {}).name || key;
                                        html += `<tr><td><strong>${escapeHtml(fieldName)}</strong></td><td>${escapeHtml(String(value))}</td></tr>`;
                                    });
                                    html += '</tbody></table>';
                                    document.getElementById(renderId).innerHTML = html;
                                } else if (format === 'key-value') {
                                    // 渲染为键值对
                                    let html = '<div class="parse-kv-list">';
                                    Object.entries(parseResult).forEach(([key, value]) => {
                                        const fieldName = (fields.find(f => f.key === key) || {}).name || key;
                                        html += `<div class="parse-kv-item"><span class="parse-kv-key">${escapeHtml(fieldName)}</span><span class="parse-kv-value">${escapeHtml(String(value))}</span></div>`;
                                    });
                                    html += '</div>';
                                    document.getElementById(renderId).innerHTML = html;
                                } else {
                                    // JSON格式
                                    document.getElementById(renderId).innerHTML = '<pre style="margin:0;padding:12px;background:white;border-radius:4px;font-size:12px;line-height:1.5;overflow:auto;">' + escapeHtml(JSON.stringify(parseResult, null, 2)) + '</pre>';
                                }
                            } else {
                                fallbackRender(body, resultText);
                            }
                        } catch (e) {
                            ChatLog.error('OCR', '[Batch] 智能解析失败:', e);
                            // 降级到普通渲染
                            fallbackRender(body, resultText);
                        }
                    } else {
                        fallbackRender(body, resultText);
                    }
                } else {
                    fallbackRender(body, resultText);
                }
            } else {
                // 未启用智能解析，使用原有渲染
                fallbackRender(body, resultText);
            }
        } else {
            card.style.borderColor = '#ffcdd2';
            header.innerHTML = '<span style="font-size:12px;font-weight:600;color:#c62828;flex:1;">' +
                escapeHtml(file.name) + '</span>' +
                '<span style="font-size:11px;color:#f44336;">❌ 失败</span>';
            body.innerHTML = '<span style="color:#d32f2f;font-size:11px;">' + escapeHtml(errorMsg || '未知错误') + '</span>';
        }

        if (_autoTrack) {
            const container = document.getElementById('batch-stream-content');
            if (container) container.scrollTop = container.scrollHeight;
        }
    }

    function fallbackRender(body, resultText) {
        const format = document.getElementById('ocr-format-select')?.value || 'plaintext';
        if (typeof OcrFormatConverter !== 'undefined') {
            const renderId = 'batch-stream-render-' + Date.now();
            body.innerHTML = '<div id="' + renderId + '"></div>';
            OcrFormatConverter.renderOcrResult(renderId, resultText, format);
        } else {
            body.innerHTML = '<pre style="white-space:pre-wrap;font-family:Consolas,monospace;font-size:11px;margin:0;">' +
                escapeHtml(resultText) + '</pre>';
        }
    }

    async function copyAllResults() {
        const texts = BatchState.files
            .filter(f => f.status === 'completed' && f.result && !isErrorResult(f.result))
            .map(f => '=== ' + f.name + ' ===\n' + f.result);

        if (texts.length === 0) {
            showBatchToast('没有可复制的结果');
            return;
        }

        // 按钮视觉反馈：找到触发按钮并显示复制状态
        const btn = document.querySelector('[onclick="OcrBatchUI.copyAllResults()"]');
        const originalHtml = btn ? btn.innerHTML : '';

        if (btn) {
            btn.innerHTML = '&#10003; 已复制';
            btn.style.background = '#4caf50';
            btn.style.color = '#fff';
            btn.style.borderColor = '#388e3c';
        }

        try {
            await navigator.clipboard.writeText(texts.join('\n\n'));
            const totalChars = texts.join('\n\n').length;
            showBatchToast(`已复制 ${texts.length} 个文件结果（${totalChars} 字符）`);
        } catch (err) {
            ChatLog.warn('OCR', '批量复制失败:', err);
            // 降级方案：创建临时 textarea 复制
            try {
                const ta = document.createElement('textarea');
                ta.value = texts.join('\n\n');
                ta.style.position = 'fixed';
                ta.style.opacity = '0';
                document.body.appendChild(ta);
                ta.select();
                document.execCommand('copy');
                document.body.removeChild(ta);
                showBatchToast(`已复制 ${texts.length} 个文件结果`);
            } catch (e) {
                showBatchToast('复制失败，请手动复制');
            }
        } finally {
            // 恢复按钮状态
            setTimeout(() => {
                if (btn) {
                    btn.innerHTML = originalHtml;
                    btn.style.background = '';
                    btn.style.color = '';
                    btn.style.borderColor = '';
                }
            }, 1800);
        }
    }

    function clearStreamResults() {
        const container = document.getElementById('batch-stream-content');
        const stream = document.getElementById('batch-result-stream');

        // 记录清除前是否有内容
        const hadContent = container && container.children.length > 0;

        if (container) container.innerHTML = '';
        if (stream) stream.style.display = 'none';
        _streamItemCount = 0;

        // 清除文件列表中的选中状态
        document.querySelectorAll('.batch-file-row').forEach(el => el.classList.remove('batch-row-selected'));
        _selectedIndex = -1;

        // 状态反馈
        if (hadContent) {
            showBatchToast('结果汇总已清空');
        }
    }

    // ============================================================
    // 辅助函数
    // ============================================================

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function showBatchToast(message) {
        const toast = document.createElement('div');
        toast.className = 'toast';
        toast.textContent = message;
        document.body.appendChild(toast);
        setTimeout(() => toast.remove(), 2500);
    }

    function formatBatchFileSize(bytes) {
        if (!bytes) return '';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    function escapeHtml(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    // 导出公共 API
    window.OcrBatchUI = {
        initBatchTab: initBatchTab,
        selectBatchFolder: selectBatchFolder,
        selectBatchFiles: selectBatchFiles,
        startBatchProcessing: startBatchProcessing,
        pauseBatch: pauseBatch,
        cancelBatch: cancelBatch,
        removeBatchFile: removeBatchFile,
        clearBatchList: clearBatchList,
        addFilesToList: addFilesToList,
        exportBatchToExcel: exportBatchToExcel,
        onClickFileRow: onClickFileRow,
        startFileDrag: startFileDrag,
        toggleAutoTrack: toggleAutoTrack,
        copyAllResults: copyAllResults,
        clearStreamResults: clearStreamResults,
        parseSingleFileWithScript: parseSingleFileWithScript,
        parseSingleFile: parseSingleFile,
        parseCompletedFiles: parseCompletedFiles,
        getState: function() { return BatchState; },
        applyLocale: function() { renderFileList(); }
    };

})();