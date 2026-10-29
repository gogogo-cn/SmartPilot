/**
 * ============================================================
 * excel-file-picker.js — Excel File Picker
 * ============================================================
 * 功能：Excel 文件选择器，提供工作簿/工作表选择、范围选取与文件浏览功能。
 *
 * 核心类/函数：
 *   - pickFile() — 选择文件
 *   - pickRange() — 选择范围
 *   - pickWorksheet() — 选择工作表
 *
 * 依赖：
 *   - app.js
 *   - excel-connector.js
 * ============================================================
 */

(function() {
    'use strict';
    
    ChatLog.info('PICKER', '[ExcelFilePicker] 📦 脚本开始加载...');
    
    const selectedFiles = [];
    // Excel 文件元数据缓存：用 SheetJS 解析得到的轻量信息（工作表名、行列数等），
    // 供 AI 上下文直接使用，避免每次提问都依赖 MCP 工具读取活动工作簿。
    const fileMetadataCache = new Map();
    
    function initialize() {
        ChatLog.info('PICKER', '[ExcelFilePicker] 🚀 开始初始化...');
        
        if (typeof window.App === 'undefined') {
            ChatLog.warn('PICKER', '[ExcelFilePicker]  App 未就绪，等待...');
            setTimeout(initialize, 200);
            return;
        }
        
        const container = document.getElementById('connection-status');
        if (!container) {
            ChatLog.error('PICKER', '[ExcelFilePicker] ❌ 找不到 #connection-status！');
            return;
        }
        
        ChatLog.info('PICKER', '[ExcelFilePicker] ✅ 找到容器');
        createWidget(container);
        ChatLog.info('PICKER', '[ExcelFilePicker] ✅ 初始化完成！');
    }
    
    function createWidget(container) {
        container.innerHTML = `
            <div id="excel-file-picker" style="display:flex;align-items:center;gap:8px;position:relative;">
                <div onclick="toggleMenu()" style="cursor:pointer;display:flex;align-items:center;gap:6px;padding:4px 10px;border-radius:20px;background:rgba(255,255,255,0.15);">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                        <polyline points="14 2 14 8 20 8"></polyline>
                        <line x1="16" y1="13" x2="8" y2="13"></line>
                        <line x1="16" y1="17" x2="8" y2="17"></line>
                    </svg>
                    <span id="file-count" style="font-size:13px;font-weight:500;color:white;">0个文件</span>
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="6 9 12 15 18 9"></polyline>
                    </svg>
                </div>
                
                <div id="file-picker-menu" style="display:none;position:absolute;top:100%;right:0;margin-top:8px;background:white;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,0.15);padding:16px;min-width:420px;max-width:500px;max-height:500px;overflow-y:auto;z-index:10000;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;padding-bottom:12px;border-bottom:2px solid #f0f0f0;">
                        <div style="font-weight:600;font-size:15px;color:#333;">📊 Excel 文件管理</div>
                        <button onclick="closeMenu()" style="background:none;border:none;font-size:20px;cursor:pointer;color:#999;">&times;</button>
                    </div>
                    
                    <!-- 文件选择器 -->
                    <div style="margin-bottom:14px;">
                        <div style="font-weight:600;font-size:13px;color:#333;margin-bottom:8px;">📁 选择Excel文件（可多选）</div>
                        <input type="file" id="excel-file-input" accept=".xlsx,.xls,.xlsm" multiple style="display:none;">
                        <button id="btn-add-excel" style="width:100%;padding:12px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;border:none;border-radius:8px;cursor:pointer;font-size:14px;font-weight:600;">
                            ➕ 添加Excel文件
                        </button>
                    </div>
                    
                    <!-- 手动选择的文件列表 -->
                    <div id="selected-files-container" style="margin-bottom:14px;">
                        <div style="font-weight:600;font-size:13px;color:#333;margin-bottom:8px;">📋 已选择的文件 (<span id="selected-count">0</span>)</div>
                        <div id="selected-files-list" style="background:#f9f9f9;border-radius:8px;padding:8px;min-height:60px;max-height:200px;overflow-y:auto;">
                            <div style="text-align:center;color:#999;padding:20px;font-size:13px;">
                                暂无文件，请点击上方按钮添加
                            </div>
                        </div>
                    </div>
                    
                    <div style="display:flex;gap:8px;">
                        <button onclick="clearAllFiles()" style="flex:1;padding:10px;background:#fff;color:#e74c3c;border:1px solid #e74c3c;border-radius:8px;cursor:pointer;font-size:13px;">️ 清空</button>
                    </div>
                </div>
            </div>
        `;
        
        injectStyles();
        bindEvents();
    }
    
    function injectStyles() {
        if(document.getElementById('efp-styles')) return;
        const style = document.createElement('style');
        style.id = 'efp-styles';
        style.textContent = `
            .file-item {
                display: flex;
                align-items: center;
                padding: 8px;
                margin-bottom: 6px;
                background: white;
                border-radius: 6px;
                border: 1px solid #e0e0e0;
            }
            .file-item:hover {
                border-color: #667eea;
            }
            .file-icon {
                font-size: 20px;
                margin-right: 10px;
            }
            .file-info {
                flex: 1;
                min-width: 0;
            }
            .file-name {
                font-size: 13px;
                font-weight: 500;
                color: #333;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .file-size {
                font-size: 11px;
                color: #999;
                margin-top: 2px;
            }
            .file-remove {
                background: none;
                border: none;
                color: #e74c3c;
                cursor: pointer;
                font-size: 16px;
                padding: 4px 8px;
                border-radius: 4px;
            }
            .file-remove:hover {
                background: #ffebee;
            }
        `;
        document.head.appendChild(style);
    }
    
    function bindEvents() {
        document.addEventListener('click', function(e) {
            const widget = document.getElementById('excel-file-picker');
            const menu = document.getElementById('file-picker-menu');
            if (widget && menu && !widget.contains(e.target)) {
                menu.style.display = 'none';
            }
        });
        
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') {
                const menu = document.getElementById('file-picker-menu');
                if (menu && menu.style.display === 'block') {
                    menu.style.display = 'none';
                }
            }
        });
        
        // 添加文件按钮点击事件
        const addBtn = document.getElementById('btn-add-excel');
        if (addBtn) {
            addBtn.addEventListener('click', function() {
                const fileInput = document.getElementById('excel-file-input');
                if (fileInput) fileInput.click();
            });
        }
        
        // 文件选择事件
        const fileInput = document.getElementById('excel-file-input');
        if (fileInput) {
            fileInput.addEventListener('change', async function(event) {
                const files = Array.from(event.target.files);
                if (files.length === 0) return;
                
                for (const file of files) {
                    if (!selectedFiles.find(f => f.file.name === file.name && f.file.size === file.size)) {
                        // 【关键】尝试通过 VSTO 桥接获取完整路径，供 MCP 工具按路径打开工作簿
                        let fullPath = '';
                        try {
                            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                                const pathResult = await window.chrome.webview.hostObjects.excelBridge.GetWorkbookPath(file.name);
                                fullPath = pathResult ? pathResult.toString() : '';
                            }
                        } catch (e) {
                            ChatLog.warn('PICKER', `[文件选择器] 获取 ${file.name} 路径失败:`, e.message);
                        }

                        const fileEntry = { file, path: fullPath };
                        selectedFiles.push(fileEntry);
                        ChatLog.info('PICKER', `[文件选择器] 添加文件: ${file.name}, 路径: ${fullPath || '(未找到，文件可能未在 Excel 中打开)'}`);

                        // 异步解析 Excel 元数据，失败不影响主流程
                        parseExcelFile(file).then(meta => {
                            fileMetadataCache.set(file.name, meta);
                            ChatLog.info('PICKER', `[文件选择器] 解析完成: ${file.name}, 工作表数: ${meta.sheetNames.length}`);
                        }).catch(e => {
                            ChatLog.warn('PICKER', `[文件选择器] 解析 ${file.name} 失败:`, e.message);
                        });
                    }
                }

                syncToAppSelectedFiles();
                renderSelectedFiles();
                updateFileCount();
                event.target.value = '';
            });
        }
    }
    
    window.toggleMenu = function() {
        const menu = document.getElementById('file-picker-menu');
        if (menu) {
            menu.style.display = menu.style.display === 'block' ? 'none' : 'block';
        }
    };
    
    window.closeMenu = function() {
        const menu = document.getElementById('file-picker-menu');
        if (menu) menu.style.display = 'none';
    };
    
    window.renderSelectedFiles = function() {
        const list = document.getElementById('selected-files-list');
        const count = document.getElementById('selected-count');
        
        if (!list) return;
        
        count.textContent = selectedFiles.length;
        
        if (selectedFiles.length === 0) {
            list.innerHTML = '<div style="text-align:center;color:#999;padding:20px;font-size:13px;">暂无文件，请点击上方按钮添加</div>';
            return;
        }
        
        list.innerHTML = selectedFiles.map((entry, index) => `
            <div class="file-item">
                <span class="file-icon"></span>
                <div class="file-info">
                    <div class="file-name" title="${entry.file.name}">${entry.file.name}</div>
                    <div class="file-size">${formatFileSize(entry.file.size)}</div>
                </div>
                <button class="file-remove" onclick="removeFile(${index})" title="移除">✕</button>
            </div>
        `).join('');
    };
    
    window.removeFile = function(index) {
        const removed = selectedFiles[index];
        if (removed) {
            fileMetadataCache.delete(removed.file.name);
        }
        selectedFiles.splice(index, 1);
        syncToAppSelectedFiles();
        renderSelectedFiles();
        updateFileCount();
    };

    window.clearAllFiles = function() {
        if (selectedFiles.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">没有可清空的文件</span>', true) : showToast('没有可清空的文件', 'info');
            return;
        }

        if (confirm('确定清空所有文件？')) {
            selectedFiles.length = 0;
            fileMetadataCache.clear();
            syncToAppSelectedFiles();
            renderSelectedFiles();
            updateFileCount();
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已清空所有文件</span>', true) : showToast('已清空所有文件', 'success');
        }
    };

    // 将 Excel 文件选择器的选中文件同步到 App.selectedFiles，
    // 让 app-chat-ui.js 等通用流程也能识别到这些附件。
    function syncToAppSelectedFiles() {
        if (typeof window.App !== 'undefined') {
            App.selectedFiles = selectedFiles.map(entry => entry.file);
            App.selectedFile = null; // 多文件模式下清空单文件引用，避免歧义
            window.selectedFiles = App.selectedFiles;
            window.selectedFile = null;
            window.__backupSelectedFiles = App.selectedFiles;
            window.__backupSelectedFile = null;
            ChatLog.info('PICKER', '[ExcelFilePicker] 已同步到 App.selectedFiles:', App.selectedFiles.length);
        }
    }
    
    window.formatFileSize = function(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
    };
    
    window.updateFileCount = function() {
        const countEl = document.getElementById('file-count');
        if (countEl) {
            countEl.textContent = selectedFiles.length + '个文件';
        }
    };
    
    window.showToast = function(msg, type) {
        const toast = document.createElement('div');
        const colors = {
            success: '#4caf50',
            warning: '#ff9800',
            error: '#f44336',
            info: '#2196f3'
        };
        toast.style.cssText = `position:fixed;top:20px;right:20px;background:${colors[type] || colors.info};color:white;padding:14px 24px;border-radius:10px;box-shadow:0 4px 20px rgba(0,0,0,0.2);z-index:10002;font-size:14px;`;
        toast.textContent = msg;
        document.body.appendChild(toast);
        setTimeout(() => toast.remove(), 3000);
    };
    
    // 暴露给App命名空间
    window.App.ExcelFiles = {
        getSelectedFiles: () => selectedFiles,
        getAllFiles: () => selectedFiles,

        // 获取已选 Excel 文件的轻量元数据（工作表名、行列数等），供 AI 上下文使用。
        // 注意：只返回元数据，不返回单元格内容，符合"AI 通过 MCP 工具实时读取 Excel 数据"的设计。
        getExcelFileSummaries: () => {
            const summaries = [];
            for (const entry of selectedFiles) {
                const meta = fileMetadataCache.get(entry.file.name);
                summaries.push({
                    name: entry.file.name,
                    size: entry.file.size,
                    path: entry.path,
                    sheetCount: meta ? meta.sheetNames.length : null,
                    sheetNames: meta ? meta.sheetNames : null,
                    sheets: meta ? meta.sheets : null
                });
            }
            return summaries;
        },

        // 获取文件内容（用于传递给AI）- Base64备用
        getFileContents: async () => {
            const contents = [];
            
            for (const entry of selectedFiles) {
                try {
                    const content = await readFileAsBase64(entry.file);
                    contents.push({
                        name: entry.file.name,
                        size: entry.file.size,
                        path: entry.path,
                        type: 'file',
                        data: content
                    });
                } catch (e) {
                    ChatLog.error('PICKER', `读取文件 ${entry.file.name} 失败:`, e);
                }
            }
            
            return contents;
        }
    };
    
    // 解析Excel文件为结构化数据
    async function parseExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 提取工作簿信息
                    const result = {
                        name: file.name,
                        size: file.size,
                        sheetNames: workbook.SheetNames,
                        sheets: {}
                    };
                    
                    // 提取每个工作表的数据（前50行）
                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        // 只保留前50行，避免数据过大
                        result.sheets[sheetName] = {
                            totalRows: jsonData.length,
                            totalCols: jsonData.length > 0 ? jsonData[0].length : 0,
                            preview: jsonData.slice(0, 50)  // 前50行预览
                        };
                    });
                    
                    resolve(result);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('文件读取失败'));
            reader.readAsArrayBuffer(file);
        });
    }
    
    // 读取文件为Base64
    function readFileAsBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }
    
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initialize);
    } else {
        initialize();
    }
})();
