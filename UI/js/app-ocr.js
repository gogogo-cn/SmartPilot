/**
 * ============================================================
 * app-ocr.js — OCR Engine Manager
 * ============================================================
 * 功能：管理 PaddleOCR / Windows OCR / Cloud OCR 三种 OCR 引擎的安装、
 *       检测、配置与调用，提供统一的图片文字识别接口。
 *
 * 核心函数：
 *   - recognizeImage(imagePath, options) — 识别图片文字
 *   - installOcrEngine(engineName) — 安装指定引擎
 *   - detectAvailableEngines() — 检测可用引擎
 *   - getCurrentOcrEngine() — 获取当前引擎
 *   - saveOcrSettings(settings) — 保存 OCR 设置
 *
 * 依赖：
 *   - app.js（App.config、App.showToast）
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    // --- 来自 app.js:L3610-L3611 ---
    let isInstallingOcr = false;

    // --- 来自 app.js:L3613-L3670 ---
    // Tesseract 安装状态检查已移除（Tesseract 已被从项目中彻底删除）
    async function checkOcrStatus() {
        // 此函数保留为空操作以兼容旧调用点
    }

    // ==================== OCR 引擎管理 ====================

    async function installOcrEngine() {
        if (isInstallingOcr) return;

        const installBtn = document.getElementById('btn-install-ocr');
        const progressDiv = document.getElementById('ocr-progress');
        const progressBar = document.getElementById('ocr-progress-bar');
        const progressText = document.getElementById('ocr-progress-text');
        const statusDiv = document.getElementById('ocr-status');
        const manualSection = document.getElementById('manual-install-section');

        // 清空之前的日志
        clearOcrLog();

        isInstallingOcr = true;
        installBtn.disabled = true;
        installBtn.textContent = uiText('ocr.status.installing');
        progressDiv.style.display = 'block';

        // 隐藏手动安装区域
        if (manualSection) {
            manualSection.style.display = 'none';
        }

        addOcrLog('开始安装 OCR 引擎...');

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                addOcrLog('调用后端安装程序...');
                // 开始安装
                const resultJson = await window.chrome.webview.hostObjects.excelBridge.InstallOcrEngine();
                addOcrLog('收到后端响应');
                const result = JSON.parse(resultJson);

                if (result.success) {
                    progressBar.style.width = '100%';
                    const mirrorInfo = result.usedMirror ? ` (使用${result.usedMirror})` : '';
                    progressText.textContent = '安装完成！' + mirrorInfo;
                    statusDiv.innerHTML = '✅ OCR 引擎已安装' + mirrorInfo;
                    statusDiv.style.color = '#4caf50';
                    installBtn.style.display = 'none';
                    showToast('OCR 引擎安装成功！' + mirrorInfo);
                } else {
                    const mirrorInfo = result.usedMirror ? ` (尝试${result.usedMirror}失败)` : '';
                    throw new Error((result.error || '安装失败') + mirrorInfo);
                }
            } else {
                throw new Error('WebView2 未就绪');
            }
        } catch (ex) {
            ChatLog.error('OCR', '安装 OCR 失败:', ex);
            progressText.textContent = '自动安装失败，请使用手动安装';
            progressBar.style.background = '#f44336';
            showToast('自动安装失败，请使用手动安装方式');

            // 显示手动安装选项
            const manualSection = document.getElementById('manual-install-section');
            if (manualSection) {
                manualSection.style.display = 'block';
            }

            installBtn.disabled = false;
            installBtn.textContent = '重试自动安装';
        } finally {
            isInstallingOcr = false;
        }
    }

    function updateOcrProgress(percent, message) {
        const progressBar = document.getElementById('ocr-progress-bar');
        const progressText = document.getElementById('ocr-progress-text');

        if (progressBar && progressText) {
            progressBar.style.width = percent + '%';
            progressText.textContent = message;
        }

        // 同时添加到日志区域
        addOcrLog(message);
    }

    function addOcrLog(message) {
        const logSection = document.getElementById('ocr-log-section');
        const logContent = document.getElementById('ocr-log-content');

        if (logSection && logContent) {
            logSection.style.display = 'block';
            const time = new Date().toLocaleTimeString();
            logContent.innerHTML += `[${time}] ${message}<br>`;
            logContent.scrollTop = logContent.scrollHeight;
        }
    }


    function clearOcrLog() {
        const logContent = document.getElementById('ocr-log-content');
        if (logContent) {
            logContent.innerHTML = '';
        }
    }

    // ==================== OCR 引擎管理 ====================


    function uiText(key, vars) {
        try {
            if (window.App && App.i18n && App.i18n.t) {
                var v = App.i18n.t(key, vars);
                if (v && v !== key) return v;
            }
        } catch (_) {}
        return key;
    }

    function uiStatus(message) {
        try {
            if (window.App && App.i18n && App.i18n.localizeOcrMessage) {
                return App.i18n.localizeOcrMessage(message);
            }
        } catch (_) {}
        return message || '';
    }

    function setStatusLine(content, color, label, detail) {
        content.textContent = '';
        var span = document.createElement('span');
        span.style.color = color;
        span.textContent = label;
        content.appendChild(span);
        if (detail) content.appendChild(document.createTextNode(' - ' + detail));
    }

    const ocrEngineDescKeys = {
        'rapidocr': 'ocr.desc.rapid',
        'paddleocr': 'ocr.desc.paddle',
        'cloud': 'ocr.desc.cloud'
    };


    // 切换 OCR 引擎
    function onOcrEngineChange() {
        const select = document.getElementById('ocr-engine-select');
        // OCR 设置面板元素可能不存在
        if (!select) { ChatLog.warn('OCR', '[onOcrEngineChange] ocr-engine-select 元素不存在，跳过'); return; }
        const engine = select.value;
        const description = document.getElementById('ocr-engine-description');

        // 更新描述
        if (description) description.textContent = uiText(ocrEngineDescKeys[engine] || 'ocr.desc.default');

        // 显示对应的设置面板
        document.querySelectorAll('.ocr-engine-panel').forEach(panel => {
            panel.style.display = 'none';
        });
        document.getElementById(engine + '-settings').style.display = 'block';

        // 更新引擎状态显示
        updateOcrEngineStatus(engine);

        // 加载各引擎的特定配置
        if (engine === 'paddleocr') {
            checkPaddleOcrStatus();
            loadPaddleOcrAdvancedConfig();
        }
    }


    // 更新 OCR 引擎状态显示
    async function updateOcrEngineStatus(engine) {
        const statusContent = document.getElementById('ocr-engine-status-content');
        if (!statusContent) return;

        // 验证引擎类型是否有效
        const validEngines = ['rapidocr', 'paddleocr', 'windows', 'cloud'];
        if (!engine || !validEngines.includes(engine)) {
            setStatusLine(statusContent, '#ff9800', uiText('ocr.status.notSelected'), uiText('ocr.status.pickEngine'));
            return;
        }

        setStatusLine(statusContent, '#2196f3', uiText('ocr.status.checking'));

        try {
            const status = await callOcrBridge('GetOcrEngineStatus', [engine]);
            const statusObj = JSON.parse(status);

            if (statusObj.available) {
                setStatusLine(statusContent, '#4caf50', uiText('ocr.status.available'), uiStatus(statusObj.message) || uiText('ocr.status.ready'));
                if (engine === 'rapidocr') {
                    const rapidStatusDiv = document.getElementById('rapidocr-status');
                    if (rapidStatusDiv) {
                        rapidStatusDiv.textContent = uiText('ocr.status.rapidOk');
                        rapidStatusDiv.style.color = '#4caf50';
                        rapidStatusDiv.style.background = '#e8f5e9';
                    }
                }
            } else {
                setStatusLine(statusContent, '#ff9800', uiText('ocr.status.unavailable'), uiStatus(statusObj.message) || uiText('ocr.status.needConfig'));
                if (engine === 'rapidocr') {
                    const rapidStatusDiv = document.getElementById('rapidocr-status');
                    if (rapidStatusDiv) {
                        rapidStatusDiv.textContent = uiText('ocr.status.rapidBad') + ': ' + (uiStatus(statusObj.message) || uiText('ocr.status.checkPlugin'));
                        rapidStatusDiv.style.color = '#ff9800';
                        rapidStatusDiv.style.background = '#fff3e0';
                    }
                }
            }
        } catch (ex) {
            ChatLog.error('OCR', '获取 OCR 引擎状态失败:', ex);
            setStatusLine(statusContent, '#f44336', uiText('ocr.status.fail'), uiText('ocr.status.checkBackend'));
        }
    }

    // --- RapidOCR 状态检查 ---
    async function checkRapidOcrStatus() {
        const statusDiv = document.getElementById('rapidocr-status');

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                // 强制重新检查（清除缓存）
                const status = await window.chrome.webview.hostObjects.excelBridge.ForceCheckRapidOcrStatus();
                const statusObj = JSON.parse(status);

                if (statusObj.available) {
                    statusDiv.textContent = uiText('ocr.status.rapidOk');
                    statusDiv.style.color = '#4caf50';
                    statusDiv.style.background = '#e8f5e9';
                } else {
                    statusDiv.textContent = uiText('ocr.status.rapidBad') + ': ' + (uiStatus(statusObj.message) || uiText('ocr.status.checkPlugin'));
                    statusDiv.style.color = '#ff9800';
                    statusDiv.style.background = '#fff3e0';
                }
            } else {
                statusDiv.textContent = uiText('ocr.status.waitInit');
            }
        } catch (ex) {
            ChatLog.error('OCR', '检查 RapidOCR 状态失败:', ex);
            statusDiv.textContent = uiText('ocr.status.fail') + ': ' + ex.message;
            statusDiv.style.color = '#f44336';
            statusDiv.style.background = '#ffebee';
        }
    }


    // 检查 PaddleOCR 状态
    async function checkPaddleOcrStatus() {
        const statusDiv = document.getElementById('paddleocr-status');
        const pythonMissingDiv = document.getElementById('paddleocr-python-missing');
        const installBtn = document.getElementById('btn-install-paddleocr');
        const diagnoseBtn = document.getElementById('btn-diagnose-paddleocr');
        const oneclickInstallDiv = document.getElementById('paddleocr-oneclick-install');
        const manualInstallDiv = document.getElementById('paddleocr-manual-install');

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const status = await window.chrome.webview.hostObjects.excelBridge.GetPaddleOcrStatus();
                const statusObj = JSON.parse(status);

                if (statusObj.available) {
                    statusDiv.textContent = uiText('ocr.paddle.ok');
                    statusDiv.style.color = '#4caf50';
                    pythonMissingDiv.style.display = 'none';
                    installBtn.style.display = 'none';
                    diagnoseBtn.style.display = 'inline-block';
                    if (oneclickInstallDiv) oneclickInstallDiv.style.display = 'none';
                    if (manualInstallDiv) manualInstallDiv.style.display = 'none';
                } else if (statusObj.python_found === 'False') {
                    statusDiv.textContent = uiText('ocr.status.noPython');
                    statusDiv.style.color = '#ff9800';
                    pythonMissingDiv.style.display = 'block';
                    installBtn.style.display = 'none';
                    diagnoseBtn.style.display = 'inline-block';
                    if (oneclickInstallDiv) oneclickInstallDiv.style.display = 'block';
                    if (manualInstallDiv) manualInstallDiv.style.display = 'block';
                } else {
                    statusDiv.textContent = uiText('ocr.paddle.notInstalled');
                    statusDiv.style.color = '#ff9800';
                    pythonMissingDiv.style.display = 'none';
                    installBtn.style.display = 'inline-block';
                    diagnoseBtn.style.display = 'inline-block';
                    if (oneclickInstallDiv) oneclickInstallDiv.style.display = 'block';
                    if (manualInstallDiv) manualInstallDiv.style.display = 'block';
                }
            } else {
                statusDiv.textContent = uiText('ocr.paddle.cannotCheck');
            }
        } catch (ex) {
            ChatLog.error('OCR', '检查 PaddleOCR 状态失败:', ex);
            statusDiv.textContent = uiText('ocr.status.fail');
        }
    }


    // 安装 PaddleOCR（后台启动 + 轮询实时进度）
    let _paddleOcrPollInterval = null;

    async function installPaddleOcr() {
        const installBtn = document.getElementById('btn-install-paddleocr');
        const progressDiv = document.getElementById('paddleocr-progress');
        const progressBar = document.getElementById('paddleocr-progress-bar');
        const progressText = document.getElementById('paddleocr-progress-text');
        const logSection = document.getElementById('paddleocr-log-section');
        const logContent = document.getElementById('paddleocr-log-content');

        if (_paddleOcrPollInterval) {
            clearInterval(_paddleOcrPollInterval);
            _paddleOcrPollInterval = null;
        }

        installBtn.disabled = true;
        installBtn.textContent = uiText('ocr.status.installing');
        progressDiv.style.display = 'block';
        logSection.style.display = 'block';
        logContent.innerHTML = '';
        progressBar.style.width = '10%';
        progressBar.style.background = '#2196f3';
        progressText.textContent = uiText('ocr.status.startingInstall');

        try {
            if (!window.chrome?.webview?.hostObjects?.excelBridge) {
                throw new Error('WebView2 未就绪');
            }

            // 步骤1：启动安装（立即返回）
            addPaddleOcrLog('🚀 正在启动 PaddleOCR 安装...');
            const startResult = await window.chrome.webview.hostObjects.excelBridge.InstallPaddleOcr();

            let started = false;
            try {
                const parsed = JSON.parse(startResult);
                started = parsed.success && parsed.started;
            } catch (e) { (window.ChatLog || console).debug?.('UI', '[ocr] 安装启动结果非 JSON，视为未启动:', e); }

            if (!started) {
                throw new Error('安装启动失败: ' + (startResult || '未知错误'));
            }

            addPaddleOcrLog('✅ 安装已在后台运行...');
            addPaddleOcrLog('');
            addPaddleOcrLog('⏳ 安装过程可能需要 5-15 分钟，请耐心等待...');

            // 步骤2：轮询实时进度
            let lastLineCount = 0;
            _paddleOcrPollInterval = setInterval(async () => {
                try {
                    const progressJson = await window.chrome.webview.hostObjects.excelBridge.GetPaddleOcrInstallProgress();
                    const data = JSON.parse(progressJson);
            
                    if (data.log) {
                        const lines = data.log.split('\n').filter(l => l.trim());
                        if (lines.length > lastLineCount) {
                            for (let i = lastLineCount; i < lines.length; i++) {
                                addPaddleOcrLog(lines[i].trim());
                            }
                            lastLineCount = lines.length;
            
                            const recentLog = lines.join('\n').toLowerCase();
                            let stage = '正在启动安装...';
                            let pct = 10;
            
                            if (recentLog.includes('step 4') || recentLog.includes('模型') || recentLog.includes('paddleocr(use_textline_orientation')) {
                                stage = '正在下载 OCR 识别模型（约 100-500MB）...';
                                pct = 75;
                            } else if (recentLog.includes('step 3') || recentLog.includes('安装 paddleocr')) {
                                stage = '正在安装 PaddleOCR 包...';
                                pct = 50;
                            } else if (recentLog.includes('step 2') || recentLog.includes('安装 paddlepaddle')) {
                                stage = '正在安装 PaddlePaddle 引擎...';
                                pct = 25;
                            } else if (recentLog.includes('step 1') || recentLog.includes('升级 pip') || recentLog.includes('pip 升级')) {
                                stage = '正在升级 pip...';
                                pct = 10;
                            } else if (recentLog.includes('success') && recentLog.includes('所有组件安装完成')) {
                                stage = '安装完成！';
                                pct = 95;
                            }
            
                            progressBar.style.width = pct + '%';
                            progressText.textContent = stage;
                            installBtn.textContent = stage;
                        }
                    }
            
                    // 检查是否完成
                    if (data.completed) {
                        clearInterval(_paddleOcrPollInterval);
                        _paddleOcrPollInterval = null;
                        progressBar.style.width = '100%';
            
                        if (data.success) {
                            progressBar.style.background = '#4caf50';
                            progressText.textContent = '安装完成！';
                            installBtn.textContent = '安装完成！';
                            showToast('PaddleOCR 安装成功！');
                            checkPaddleOcrStatus();
                        } else {
                            progressBar.style.background = '#f44336';
                            progressText.textContent = '安装失败';
                            installBtn.textContent = '安装失败，点击重试';
                            addPaddleOcrLog('');
                            addPaddleOcrLog('❌ 安装失败，请尝试一键安装或手动安装');
                        }
            
                        installBtn.disabled = false;
                    }
                } catch (pollErr) {
                    ChatLog.error('OCR', '轮询进度失败:', pollErr);
                }
            }, 1500); // 每1.5秒轮询一次

        } catch (ex) {
            if (_paddleOcrPollInterval) {
                clearInterval(_paddleOcrPollInterval);
                _paddleOcrPollInterval = null;
            }
            ChatLog.error('OCR', '安装失败:', ex);
            progressText.textContent = '安装失败';
            progressBar.style.background = '#f44336';
            addPaddleOcrLog('');
            addPaddleOcrLog('❌ 错误: ' + (ex.message || '安装失败'));
            installBtn.disabled = false;
            installBtn.textContent = '📥 自动安装 PaddleOCR';
        }
    }


    // 添加 PaddleOCR 安装日志
    function addPaddleOcrLog(message) {
        const logContent = document.getElementById('paddleocr-log-content');
        if (logContent) {
            const time = new Date().toLocaleTimeString();
            const line = document.createElement('div');
            line.textContent = `[${time}] ${message}`;
            logContent.appendChild(line);
            logContent.scrollTop = logContent.scrollHeight;
        }
    }

    // --- 验证/诊断 PaddleOCR 安装 ---
    async function diagnosePaddleOcr() {
        const btn = document.getElementById('btn-diagnose-paddleocr');
        const section = document.getElementById('paddleocr-diagnose-section');
        const content = document.getElementById('paddleocr-diagnose-content');

        if (!window.chrome?.webview?.hostObjects?.excelBridge) {
            showToast('WebView2 未就绪');
            return;
        }

        btn.disabled = true;
        btn.textContent = '🔍 正在诊断...';
        section.style.display = 'block';
        content.innerHTML = '<div style="color: #666;">⏳ 正在扫描所有 Python 环境，请稍候（约15-30秒）...</div>';

        try {
            const resultJson = await window.chrome.webview.hostObjects.excelBridge.DiagnosePaddleOcr();
            const data = JSON.parse(resultJson);

            let html = '';
            html += `<div style="color: #333; font-weight: bold; margin-bottom: 8px;">
                找到 ${data.total_pythons || 0} 个 Python 环境
                ${data.any_success ? '<span style="color: #4caf50;">✅ 其中至少1个支持 PaddleOCR</span>' : '<span style="color: #f44336;">❌ 无一支持 PaddleOCR</span>'}
            </div>`;
            html += '<hr style="margin: 8px 0; border-color: #ddd;">';

            if (data.details && data.details.length > 0) {
                for (const detail of data.details) {
                    const okIcon = detail.paddleocr_ok ? '✅' : '❌';
                    const pipIcon = detail.pip_paddleocr ? '✅' : '⚠️';
                    const pyIcon = detail.python_ok ? '✅' : '❌';
                    const bgColor = detail.paddleocr_ok ? '#e8f5e9' : '#fff3e0';

                    html += `<div style="background: ${bgColor}; border-radius: 4px; padding: 8px; margin-bottom: 6px; font-size: 11px;">
                        <div style="font-weight: bold; color: #333; word-break: break-all;">${okIcon} ${detail.path}</div>
                        <div style="color: #666; margin-top: 3px;">
                            Python ${pyIcon} ${detail.version || '未知'} &nbsp;|&nbsp;
                            pip paddleocr ${pipIcon} &nbsp;|&nbsp;
                            import paddleocr ${okIcon}
                        </div>`;

                    if (!detail.paddleocr_ok && detail.import_error) {
                        html += `<div style="color: #d32f2f; margin-top: 3px; font-size: 10px;">原因: ${detail.import_error}</div>`;
                    }
                    html += '</div>';
                }
            }

            if (data.error) {
                html += `<div style="color: #d32f2f;">诊断异常: ${data.error}</div>`;
            }

            html += '<hr style="margin: 8px 0; border-color: #ddd;">';
            html += '<div style="font-size: 10px; color: #999;">提示：如果 pip 显示已安装但 import 失败，可能是 Python 环境不一致</div>';

            content.innerHTML = html;

        } catch (ex) {
            content.innerHTML = `<div style="color: #d32f2f;">诊断失败: ${ex.message}</div>`;
        } finally {
            btn.disabled = false;
            btn.textContent = '🔍 验证 PaddleOCR 安装';
        }
    }

    // 运行一键安装脚本
    async function runPaddleOCRInstaller() {
        const btn = document.getElementById('btn-oneclick-install');
        const status = document.getElementById('oneclick-install-status');

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                btn.disabled = true;
                btn.textContent = '⏳ 正在安装...';
                status.style.display = 'block';
                status.textContent = '正在运行安装程序，请查看弹出的命令窗口...';

                await window.chrome.webview.hostObjects.excelBridge.RunPaddleOCRInstaller();

                showToast('安装程序已启动，请查看弹出的命令窗口');

                // 开始轮询检查安装状态
                status.textContent = '安装程序已启动，正在等待安装完成...';
                pollPaddleOcrInstallationStatus(btn, status);
            } else {
                showToast('无法启动安装程序');
            }
        } catch (ex) {
            ChatLog.error('OCR', '启动安装程序失败:', ex);
            showToast('启动安装程序失败: ' + ex.message);
            btn.disabled = false;
            btn.textContent = '🖥️ 一键安装 PaddleOCR';
            status.style.display = 'none';
        }
    }

    // 轮询检查 PaddleOCR 安装状态
    async function pollPaddleOcrInstallationStatus(btn, status) {
        const startTime = Date.now();
        const maxWait = 10 * 60 * 1000; // 最多等待 10 分钟

        const checkInterval = setInterval(async () => {
            const elapsedMs = Date.now() - startTime;
            const elapsedMin = Math.floor(elapsedMs / 60000);
            const elapsedSec = Math.floor((elapsedMs % 60000) / 1000);

            try {
                // 检查 PaddleOCR 是否已安装
                const isInstalled = await checkPaddleOcrInstalled();

                if (isInstalled) {
                    // 安装成功
                    clearInterval(checkInterval);
                    btn.textContent = '安装完成';
                    btn.style.background = '#4caf50';
                    status.textContent = 'PaddleOCR 安装成功！请重启 Excel';
                    showToast('PaddleOCR 安装成功！请重启 Excel');

                    // 更新引擎状态显示
                    updateOcrEngineStatus('paddleocr');
                    return;
                }

                if (elapsedMs >= maxWait) {
                    // 超时
                    clearInterval(checkInterval);
                    btn.textContent = '检查超时';
                    btn.disabled = false;
                    status.textContent = '安装检查超时，请查看命令窗口确认安装结果';
                    showToast('安装检查超时，如果命令窗口显示完成，请重启 Excel');
                    return;
                }

                // 继续等待 - 显示已等待时间
                btn.textContent = `正在安装... 已等待 ${elapsedMin}分${elapsedSec}秒`;

            } catch (ex) {
                ChatLog.error('OCR', '检查安装状态失败:', ex);
            }
        }, 5000); // 每5秒检查一次
    }


    // 检查 PaddleOCR 是否已安装
    async function checkPaddleOcrInstalled() {
        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const status = await window.chrome.webview.hostObjects.excelBridge.GetOcrEngineStatus('paddleocr');
                const statusObj = JSON.parse(status);
                return statusObj.available === true;
            }
        } catch (ex) {
            ChatLog.error('OCR', '检查 PaddleOCR 安装状态失败:', ex);
        }
        return false;
    }



    // 检查 Windows OCR 状态
    async function checkWindowsOcrStatus() {
        const statusDiv = document.getElementById('windows-ocr-status');
        const languagesDiv = document.getElementById('windows-ocr-languages');
        const languagesList = document.getElementById('windows-ocr-languages-list');

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const status = await window.chrome.webview.hostObjects.excelBridge.GetWindowsOcrStatus();
                const statusObj = JSON.parse(status);

                if (statusObj.available) {
                    statusDiv.innerHTML = '✅ Windows OCR 可用';
                    statusDiv.style.color = '#4caf50';

                    // 显示支持的语言
                    if (statusObj.languages && statusObj.languages.length > 0) {
                        languagesList.innerHTML = statusObj.languages.map(lang => `<div>${lang}</div>`).join('');
                        languagesDiv.style.display = 'block';
                    }
                } else {
                    statusDiv.innerHTML = '❌ Windows OCR 不可用 - 需要 Windows 10 版本 1803 或更高';
                    statusDiv.style.color = '#f44336';
                    languagesDiv.style.display = 'none';
                }
            } else {
                statusDiv.innerHTML = '⚠️ 无法检测状态';
            }
        } catch (ex) {
            ChatLog.error('OCR', '检查 Windows OCR 状态失败:', ex);
            statusDiv.innerHTML = '⚠️ 检测失败';
        }
    }

    // 切换云端 OCR 提供商
    function onCloudProviderChange() {
        const provider = document.getElementById('cloud-ocr-provider').value;

        document.querySelectorAll('.cloud-provider-config').forEach(config => {
            config.style.display = 'none';
        });
        document.getElementById(provider + '-config').style.display = 'block';
    }


    // 测试云端 OCR 配置
    async function testCloudOcrConfig() {
        const btn = document.getElementById('btn-test-cloud-ocr');
        const resultDiv = document.getElementById('cloud-ocr-test-result');
        const provider = document.getElementById('cloud-ocr-provider').value;

        // 显示按钮加载状态
        if (btn) {
            btn.classList.add('btn-loading');
            btn.disabled = true;
            btn.dataset.originalText = btn.innerHTML;
            btn.innerHTML = '⏳ 测试中';
        }

        resultDiv.style.display = 'block';
        resultDiv.innerHTML = '<div style="color: #666;">正在验证凭据...</div>';

        try {
            let config = {};
            if (provider === 'baidu') {
                config.apiKey = document.getElementById('baidu-api-key').value;
                config.secretKey = document.getElementById('baidu-secret-key').value;
            } else if (provider === 'tencent') {
                config.secretId = document.getElementById('tencent-secret-id').value;
                config.secretKey = document.getElementById('tencent-secret-key').value;
            }

            const result = await callOcrBridge('TestCloudOcrConfig', [provider, JSON.stringify(config)]);
                const resultObj = JSON.parse(result);

                // 恢复按钮状态
                if (btn) {
                    btn.classList.remove('btn-loading');
                    btn.disabled = false;
                    btn.innerHTML = btn.dataset.originalText || '🧪 测试连接';
                }

                if (resultObj.success) {
                    resultDiv.innerHTML = '<div style="color: #4caf50; padding: 10px; background: #e8f5e9; border-radius: 4px;">✅ 连接成功！凭据验证通过</div>';
                    // 按钮成功动画
                    if (btn) {
                        btn.classList.add('btn-success-state');
                        setTimeout(() => btn.classList.remove('btn-success-state'), 1000);
                    }
                    showToast('云端 OCR 连接成功！', 'success');
                } else {
                    resultDiv.innerHTML = `<div style="color: #f44336; padding: 10px; background: #ffebee; border-radius: 4px;">❌ 连接失败: ${resultObj.error || resultObj.message || '未知错误'}</div>`;
                    // 按钮失败动画
                    if (btn) {
                        btn.classList.add('btn-error-state');
                        setTimeout(() => btn.classList.remove('btn-error-state'), 500);
                    }
                    showToast('连接失败，请检查配置', 'error');
                }
        } catch (ex) {
            ChatLog.error('OCR', '测试云端 OCR 失败:', ex);
            // 恢复按钮状态
            if (btn) {
                btn.classList.remove('btn-loading');
                btn.disabled = false;
                btn.innerHTML = btn.dataset.originalText || '🧪 测试连接';
                btn.classList.add('btn-error-state');
                setTimeout(() => btn.classList.remove('btn-error-state'), 500);
            }
            resultDiv.innerHTML = `<div style="color: #f44336; padding: 10px; background: #ffebee; border-radius: 4px;">❌ 测试失败: ${ex.message}</div>`;
            showToast('测试失败: ' + ex.message, 'error');
        }
    }

    // Toast 提示函数：所有提示统一合并到固定状态栏 fsb-content 中，不再创建独立的 toast-container
    function showToast(message, type) {
        const fsb = App.fixedStatusBar;
        if (!fsb) {
            ChatLog.info('OCR', '[toast]', message);
            return;
        }
        const t = type || 'info';
        if (t === 'success' && typeof fsb.setSuccess === 'function') {
            fsb.setSuccess(message);
        } else if (t === 'error' && typeof fsb.setError === 'function') {
            fsb.setError(message);
        } else if (t === 'warning' && typeof fsb.setWarning === 'function') {
            fsb.setWarning(message);
        } else if (typeof fsb.setInfo === 'function') {
            fsb.setInfo(message);
        } else {
            fsb.setStatus('idle', message);
        }
    }

    // 保存 OCR 设置
    async function saveOcrSettings() {
        const btn = document.getElementById('btn-save-ocr-settings');
        const engine = document.getElementById('ocr-engine-select').value;

        // 显示按钮加载状态
        if (btn) {
            btn.classList.add('btn-loading');
            btn.disabled = true;
            btn.dataset.originalText = btn.innerHTML;
            btn.innerHTML = '⏳ 保存中';
        }

        try {
            await callOcrBridge('SetOcrEngine', [engine]);
            window.__currentOcrEngine = engine;   // 供 OCR 免费额度判断是否需要计数

            if (engine === 'cloud') {
                const provider = document.getElementById('cloud-ocr-provider').value;
                let config = { provider: provider };

                if (provider === 'baidu') {
                    config.apiKey = document.getElementById('baidu-api-key').value;
                    config.secretKey = document.getElementById('baidu-secret-key').value;
                } else if (provider === 'tencent') {
                    config.secretId = document.getElementById('tencent-secret-id').value;
                    config.secretKey = document.getElementById('tencent-secret-key').value;
                }

                await callOcrBridge('SaveCloudOcrConfig', [JSON.stringify(config)]);
            }

            if (engine === 'paddleocr') {
                await savePaddleOcrAdvancedConfig();
            }

            if (btn) {
                btn.classList.remove('btn-loading');
                btn.disabled = false;
                btn.innerHTML = btn.dataset.originalText || '💾 保存 OCR 设置';
                btn.classList.add('btn-success-state');
                setTimeout(() => btn.classList.remove('btn-success-state'), 1000);
            }
            showToast('OCR 设置已保存', 'success');
            updateOcrEngineStatus(engine);
        } catch (ex) {
            ChatLog.error('OCR', '保存 OCR 设置失败:', ex);
            // 恢复按钮状态
            if (btn) {
                btn.classList.remove('btn-loading');
                btn.disabled = false;
                btn.innerHTML = btn.dataset.originalText || '💾 保存 OCR 设置';
                btn.classList.add('btn-error-state');
                setTimeout(() => btn.classList.remove('btn-error-state'), 500);
            }
            showToast('保存失败: ' + ex.message, 'error');
        }
    }


    // 加载 OCR 设置
    async function loadOcrSettings() {
        // OCR 设置面板元素可能不存在（例如从 index.html 调用时）
        const select = document.getElementById('ocr-engine-select');
        if (!select) { ChatLog.debug('OCR', '[loadOcrSettings] ocr-engine-select 元素不存在，跳过（当前页面可能非设置面板）'); return; }

            try {
                const engine = await callOcrBridge('GetCurrentOcrEngine');
                if (engine) {
                    const validEngines = ['rapidocr', 'paddleocr', 'windows', 'cloud'];
                    if (validEngines.includes(engine)) {
                        select.value = engine;
                        window.__currentOcrEngine = engine;
                } else {
                    ChatLog.warn('OCR', '未知的引擎类型:', engine, '使用默认值 rapidocr');
                    select.value = 'rapidocr';
                }
            } else {
                select.value = 'rapidocr';
            }
            onOcrEngineChange();

            if (select.value === 'paddleocr') {
                await loadPaddleOcrAdvancedConfig();
            }

            const cloudConfig = await callOcrBridge('GetCloudOcrConfig');
            if (cloudConfig) {
                const config = JSON.parse(cloudConfig);
                if (config.provider) {
                    var cloudOcrEl = document.getElementById('cloud-ocr-provider');
                    if (cloudOcrEl) cloudOcrEl.value = config.provider;
                    onCloudProviderChange();
                }
                if (config.apiKey) { var el = document.getElementById('baidu-api-key'); if (el) el.value = config.apiKey; }
                if (config.secretKey) { var el = document.getElementById('baidu-secret-key'); if (el) el.value = config.secretKey; }
                if (config.secretId) { var el = document.getElementById('tencent-secret-id'); if (el) el.value = config.secretId; }
            }
        } catch (ex) {
            ChatLog.error('OCR', '加载 OCR 设置失败:', ex);
            // 出错时更新状态显示
            const statusContent = document.getElementById('ocr-engine-status-content');
            if (statusContent) {
                statusContent.innerHTML = '<span style="color: #ff9800;">⚠️ 加载失败</span> - 请刷新页面重试';
            }
        }
    }


    // 保存OCR配置到后端
    function saveOcrConfigToBackend(apiKey, secretKey) {
        ChatLog.info('OCR', 'saveOcrConfigToBackend: 开始保存配置');
        ChatLog.info('OCR', '  apiKey:', apiKey ? apiKey.substring(0, 10) + '...' : 'empty');
        ChatLog.info('OCR', '  secretKey:', secretKey ? '***' : 'empty');

        if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
            try {
                // 直接调用 HostObject 方法保存配置
                const result = window.chrome.webview.hostObjects.excelBridge.SaveOcrConfig(apiKey, secretKey);
                ChatLog.info('OCR', 'saveOcrConfigToBackend: 结果 =', result);
            } catch (e) {
                ChatLog.error('OCR', 'saveOcrConfigToBackend: 失败', e);
            }
        } else {
            ChatLog.info('OCR', 'saveOcrConfigToBackend: HostObject 不可用，使用 postMessage');
            window.chrome.webview.postMessage({
                type: 'saveOcrConfig',
                apiKey: apiKey,
                secretKey: secretKey
            });
        }
    }

    /**
     * 当前是否云 OCR 引擎（用户自带 Key → 不占免费额度）。
     * 优先用 window.__currentOcrEngine（启动时/切引擎时写入），其次读设置页下拉。
     * 取不到时返回空（按计费通道处理，保守口径）。
     */
    function currentEngineIsCloud() {
        try {
            if (window.__currentOcrEngine) return window.__currentOcrEngine === 'cloud' ? 'cloud' : '';
            var sel = document.getElementById('ocr-engine-select');
            if (sel && sel.value) return sel.value === 'cloud' ? 'cloud' : '';
        } catch (_) { /* ignore */ }
        return '';
    }

    // ==================== PaddleOCR 高级配置 ====================

    async function callOcrBridge(method, args) {
        args = args || [];
        // 本地 OCR 每日免费额度（云通道 / 已激活不计数）。详情见 js/ocr-quota.js
        var quota = window.SmartPilotOcrQuota;
        if (quota && quota.gate({ method: method, engine: currentEngineIsCloud() }) === false) {
            throw new Error('OCR_QUOTA_EXCEEDED');
        }
        try {
            var bridge = window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge;
            if (bridge && typeof bridge[method] === 'function') {
                return await bridge[method](...args);
            }
            // 同源自适应（与 ocr-panel.html 的 OCR_HOST_URL 一致）：
            // 从桌面版桥 9527 加载时直连本地桥（OcrHost 兼容协议），插件版 OcrHost 9530 / file:// 场景保持原地址。
            var ocrHostUrl = (location.port === '9527') ? location.origin : 'http://127.0.0.1:9530';
            var resp = await fetch(ocrHostUrl + '/v1/call', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ method: 'OcrBridge.' + method, args: args })
            });
            var data = await resp.json();
            if (!resp.ok || !data.ok) throw new Error((data && data.error) || 'OCR 服务不可用');
            return data.result;
        } catch (e) {
            // 调用失败：退还这次额度（只为真正启动的识别计费）
            if (quota) quota.refund({ method: method, engine: currentEngineIsCloud() });
            throw e;
        }
    }

    function fillPaddleLanguageSelect(select, languages, selected) {
        if (!select || !languages || !languages.length) return;
        select._langs = languages;
        select.innerHTML = languages.map(function (lang) {
            var id = String(lang.id || '');
            var key = 'ocr.langId.' + id;
            var translated = uiText(key);
            var name = (translated && translated !== key) ? translated : String(lang.name || id);
            return '<option value="' + id.replace(/"/g, '&quot;') + '">' + name.replace(/</g, '&lt;') + '</option>';
        }).join('');
        if (selected) select.value = selected;
    }

    function refreshOcrLocaleTexts() {
        if (document.getElementById('ocr-engine-description')) {
            onOcrEngineChange();
            try { document.title = uiText('ocr.settingsTitle'); } catch (_) {}
        }
        var lang = document.getElementById('paddleocr-language');
        if (lang && lang._langs) fillPaddleLanguageSelect(lang, lang._langs, lang.value);
    }
    window.refreshOcrLocaleTexts = refreshOcrLocaleTexts;

    // 加载 PaddleOCR 高级配置（语言、版面/表格/公式开关）
    async function loadPaddleOcrAdvancedConfig() {
        try {
            const result = await callOcrBridge('GetPaddleOcrAdvancedConfig');
            if (!result) return;
            const data = JSON.parse(result);
            if (!data.success) return;
            const langSelect = document.getElementById('paddleocr-language');
            fillPaddleLanguageSelect(langSelect, data.supportedLanguages, data.language);
            const layoutCheck = document.getElementById('paddleocr-layout');
            if (layoutCheck) layoutCheck.checked = data.enableLayoutAnalysis !== false;
            const tableCheck = document.getElementById('paddleocr-table');
            if (tableCheck) tableCheck.checked = data.enableTableRecognition !== false;
            const formulaCheck = document.getElementById('paddleocr-formula');
            if (formulaCheck) formulaCheck.checked = data.enableFormulaRecognition !== false;
        } catch (ex) {
            ChatLog.error('OCR', '加载 PaddleOCR 高级配置失败:', ex);
        }
    }

    // 保存 PaddleOCR 高级配置
    async function savePaddleOcrAdvancedConfig() {
        try {
            const langSelect = document.getElementById('paddleocr-language');
            const layoutCheck = document.getElementById('paddleocr-layout');
            const tableCheck = document.getElementById('paddleocr-table');
            const formulaCheck = document.getElementById('paddleocr-formula');

            const config = {
                language: langSelect ? langSelect.value : 'ch',
                enableLayoutAnalysis: layoutCheck ? layoutCheck.checked : true,
                enableTableRecognition: tableCheck ? tableCheck.checked : true,
                enableFormulaRecognition: formulaCheck ? formulaCheck.checked : true
            };

            await callOcrBridge('SetPaddleOcrAdvancedConfig', [JSON.stringify(config)]);
            ChatLog.info('OCR', 'PaddleOCR 高级配置已保存:', config);
        } catch (ex) {
            ChatLog.error('OCR', '保存 PaddleOCR 高级配置失败:', ex);
        }
    }

    // ==================== 模块导出 ====================
    
    // 将所有函数暴露到 App 对象上（供其他模块和 HTML 调用）
    window.App = window.App || {};
    window.App.installOcrEngine = installOcrEngine;
    window.App.updateOcrProgress = updateOcrProgress;
    window.App.addOcrLog = addOcrLog;
    window.App.clearOcrLog = clearOcrLog;
    window.App.onOcrEngineChange = onOcrEngineChange;
    window.App.updateOcrEngineStatus = updateOcrEngineStatus;
    window.App.checkRapidOcrStatus = checkRapidOcrStatus;
    window.App.checkPaddleOcrStatus = checkPaddleOcrStatus;
    window.App.installPaddleOcr = installPaddleOcr;
    window.App.runPaddleOCRInstaller = runPaddleOCRInstaller;
    window.App.pollPaddleOcrInstallationStatus = pollPaddleOcrInstallationStatus;
    window.App.checkPaddleOcrInstalled = checkPaddleOcrInstalled;
    window.App.checkWindowsOcrStatus = checkWindowsOcrStatus;
    window.App.onCloudProviderChange = onCloudProviderChange;
    window.App.testCloudOcrConfig = testCloudOcrConfig;
    window.App.saveOcrSettings = saveOcrSettings;
    window.App.loadOcrSettings = loadOcrSettings;
    window.App.saveOcrConfigToBackend = saveOcrConfigToBackend;
    window.App.addPaddleOcrLog = addPaddleOcrLog;
    window.App.diagnosePaddleOcr = diagnosePaddleOcr;
    window.App.loadPaddleOcrAdvancedConfig = loadPaddleOcrAdvancedConfig;
    window.App.savePaddleOcrAdvancedConfig = savePaddleOcrAdvancedConfig;
    window.App.showToast = showToast;
    // 同时暴露到全局（供 HTML onclick 直接调用，防止作用域问题）
    window.onOcrEngineChange = onOcrEngineChange;
    window.checkRapidOcrStatus = checkRapidOcrStatus;
    window.checkPaddleOcrStatus = checkPaddleOcrStatus;
    window.installPaddleOcr = installPaddleOcr;
    window.testCloudOcrConfig = testCloudOcrConfig;
    window.saveOcrSettings = saveOcrSettings;
    window.diagnosePaddleOcr = diagnosePaddleOcr;
    window.showToast = showToast;

})();
