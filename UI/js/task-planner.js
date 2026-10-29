/**
 * task-planner.js - 任务规划插件（v2026082956）
 * 
 * 职责：管理定时任务的创建、编辑、删除和执行
 * 可用于 Excel 内部面板和独立页面
 * 
 * 通信方式：
 * - Excel 内部: 通过 window.taskScheduler 调用 TaskSchedulerBridge COM 接口
 * - 独立页面: 通过 HTTP API (localhost:9528)
 */
(function() {
    'use strict';
    if (window.TaskPlanner) return; // 防止重复加载

    const API_BASE = App.Endpoints ? App.Endpoints.excelVstoHttp.baseUrl + '/api' : 'http://localhost:9528/api';

    // ===== Excel 9528 API session token（/api/* 鉴权；9529 独立服务无需 token）=====
    // 优先 COM 桥（内嵌 WebView2），其次公开端点 /api/status；缓存后复用
    let _excelToken = '';
    async function getExcelToken() {
        if (_excelToken) return _excelToken;
        // 【守卫修正·B】VSTO(WebView2) 已走 COM 桥取 token，本地 9528 不存在，跳过 HTTP 探测。
        if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects) return '';
        try {
            if (window.ExcelBridge && typeof window.ExcelBridge.GetMobileControlStatus === 'function') {
                const r = window.ExcelBridge.GetMobileControlStatus();
                if (r) { const j = JSON.parse(r); if (j && j.sessionToken) { _excelToken = j.sessionToken; return _excelToken; } }
            }
        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
        try {
            const c = new AbortController(); const to = setTimeout(() => c.abort(), 1500);
            const r = await fetch(API_BASE + '/status', { cache: 'no-store', signal: c.signal });
            clearTimeout(to);
            if (r.ok) { const j = await r.json(); if (j && j.sessionToken) { _excelToken = j.sessionToken; return _excelToken; } }
        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
        return '';
    }
    function excelApiUrl(url) {
        return url + (url.indexOf('?') >= 0 ? '&' : '?') + 'token=' + encodeURIComponent(_excelToken || '');
    }

    const TaskPlanner = {
        _initialized: false,
        _tasks: [],
        _editingId: null,
        _listElId: 'tasks-list',

        /**
         * 初始化
         */
        init() {
            if (this._initialized) return;
            this._initialized = true;
            this._listElId = 'tasks-list';
            this._initAutoBypassToggle();
            this.loadTasks();
        },

        /**
         * 自动化免确认开关 - localStorage 键
         */
        _AUTO_BYPASS_KEY: 'tp_autoExecBypass',

        /**
         * 初始化自动化免确认开关（读取 localStorage 记忆状态，同步到后端与全局守卫）
         */
        _initAutoBypassToggle() {
            const toggle = document.getElementById('tp-autoexec-toggle');
            let enabled = false;
            try {
                enabled = localStorage.getItem(this._AUTO_BYPASS_KEY) === '1';
            } catch (e) { /* ignore */ }
            if (toggle) toggle.checked = enabled;
            // 全局守卫标志：tool-registry 的 dangerous-tools 守卫读取
            window._tpAutoBypass = enabled;
            this._syncAutoBypass(enabled);
        },

        /**
         * 切换自动化免确认开关（由面板 checkbox onchange 调用）
         */
        toggleAutoBypass(enabled) {
            try {
                localStorage.setItem(this._AUTO_BYPASS_KEY, enabled ? '1' : '0');
            } catch (e) { /* ignore */ }
            window._tpAutoBypass = enabled;
            this._syncAutoBypass(enabled);
            this._toast(enabled ? '已开启自动化免确认，高风险代码将不再弹确认框' : '已关闭自动化免确认', 'success');
        },

        /**
         * 同步免确认状态到后端 ExcelBridge（独立页面无 COM 桥时静默跳过）
         */
        _syncAutoBypass(enabled) {
            const bridge = window.chrome?.webview?.hostObjects?.excelBridge;
            if (bridge && typeof bridge.SetAutoExecutionBypass === 'function') {
                try {
                    bridge.SetAutoExecutionBypass(enabled);
                } catch (e) {
                    console.warn('[TaskPlanner] SetAutoExecutionBypass failed:', e.message);
                }
            }
        },

        /**
         * 从后端加载任务列表
         */
        async loadTasks() {
            const listEl = document.getElementById(this._listElId) || document.getElementById('tasks-list');
            if (!listEl) return;

            // 0. 加载前先去重（按 Name+Frequency+Time 指纹清理历史残留的重复项）
            this._deduplicate();

            // 1. 优先从 localStorage 加载（同步，立即生效）
            const local = this._loadLocal();
            if (local && local.length > 0) {
                this._tasks = local;
                this._renderTasks();
            } else {
                this._tasks = [];
                listEl.innerHTML = '<div style="text-align:center;padding:40px 20px;color:#999;font-size:13px;">暂无定时任务<br><span style="font-size:12px;color:#bbb;">点击"新建任务"按钮创建第一个定时任务</span></div>';
                this._renderStats();
            }

            // 2. 后台尝试从后端同步（失败不影响界面，空数据不覆盖本地）
            let backendTasks = [];
            try {
                if (window.taskScheduler) {
                    const json = window.taskScheduler.GetAllTasks();
                    if (json) {
                        const tasks = JSON.parse(json);
                        if (Array.isArray(tasks) && tasks.length > 0) {
                            backendTasks = tasks;
                        }
                    }
                }
            } catch (e) { /* COM Bridge not available */ }

            if (backendTasks.length === 0) {
                // Word/PPT VSTO 任务窗格没有 Excel 内嵌 9528，跳过 HTTP 探测（避免 ERR_CONNECTION_REFUSED 刷屏）
                const _isVstoHost = !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects);
                if (_isVstoHost) {
                    /* keep localStorage tasks only */
                } else try {
                    const ctrl = new AbortController();
                    const t = setTimeout(() => ctrl.abort(), 3000);
                    await getExcelToken();
                    const res = await fetch(excelApiUrl(API_BASE + '/tasks'), { signal: ctrl.signal });
                    clearTimeout(t);
                    if (res.ok) {
                        const data = await res.json();
                        const tasks = Array.isArray(data) ? data : (data?.data || []);
                        if (Array.isArray(tasks) && tasks.length > 0) {
                            backendTasks = tasks;
                        }
                    }
                } catch (e) { /* HTTP API not available */ }
            }

            // 3. 如果后端有数据，以后端为准（覆盖本地）
            if (backendTasks.length > 0) {
                this._tasks = backendTasks;
                this._persistLocal();
                this._renderTasks();
                return;
            }

            // 4. 后端无数据但本地有 → 同步到 COM Bridge（Excel 内）
            // 注意：不向 HTTP API 同步，submitInlineForm 已负责该操作，避免重复创建
            if (local && local.length > 0) {
                if (window.taskScheduler) {
                    for (const task of local) {
                        try {
                            window.taskScheduler.CreateTask(JSON.stringify(task));
                        } catch (e) { /* ignore individual failures */ }
                    }
                }
            }
        },

        /**
         * localStorage 存储键
         */
        _STORAGE_KEY: 'tp_tasks',

        /**
         * 保存到 localStorage
         */
        _persistLocal() {
            try {
                localStorage.setItem(this._STORAGE_KEY, JSON.stringify(this._tasks));
            } catch (e) {
                console.warn('[TaskPlanner] localStorage save failed:', e.message);
            }
        },

        /**
         * 从 localStorage 加载
         * @returns {Array|null}
         */
        _loadLocal() {
            try {
                const raw = localStorage.getItem(this._STORAGE_KEY);
                if (raw) return JSON.parse(raw);
            } catch (e) {
                console.warn('[TaskPlanner] localStorage load failed:', e.message);
            }
            return null;
        },

        /**
         * 生成临时 ID
         */
        _genId() {
            return 'local_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
        },

        /**
         * 前端重算下次执行时间（与 VB TaskSchedulerBridge.CalculateNextRun 行为对齐）
         * - 保证“保存后立即刷新下次执行”，无需等待后端回读
         * - 已暂停(Enabled=false) 或 仅一次且已执行过 → 返回空串
         * @param {{Frequency?:string,Time?:string,Enabled?:boolean,LastRun?:string}|null|undefined} task
         * @returns {string} "yyyy-MM-dd HH:mm:ss" 或 ""
         */
        calcNextRun(task) {
            if (!task) return '';
            if (task.Enabled === false) return '';
            if (task.Frequency === 'once' && task.LastRun) return '';

            const freq = task.Frequency || 'daily';
            const timeStr = task.Time || '09:00';

            // 每小时：对齐到下一个整点分钟；若分钟已过则下一小时
            if (freq === 'hourly') {
                const parts = String(timeStr).split(':');
                const minute = (parts.length >= 2 && /^\d+$/.test(parts[1])) ? parseInt(parts[1], 10) : 0;
                const now = new Date();
                let next = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), minute, 0, 0);
                if (next.getTime() <= now.getTime()) next = new Date(next.getTime() + 3600 * 1000);
                return this._fmtYmdHms(next);
            }

            // daily / once / weekdays / weekly：解析时+分钟，跳过到下个有效日
            const parts = String(timeStr).split(':');
            const hour = (parts.length >= 1 && /^\d+$/.test(parts[0])) ? parseInt(parts[0], 10) : 9;
            const minute = (parts.length >= 2 && /^\d+$/.test(parts[1])) ? parseInt(parts[1], 10) : 0;

            const now = new Date();
            let next = new Date(now.getFullYear(), now.getMonth(), now.getDate(), hour, minute, 0, 0);
            // 今天的时刻已过 → 推到明天再判定
            if (next.getTime() <= now.getTime()) next = new Date(next.getTime() + 24 * 3600 * 1000);

            if (freq === 'weekdays') {
                // 跳过周六周日
                while (next.getDay() === 0 || next.getDay() === 6) {
                    next = new Date(next.getTime() + 24 * 3600 * 1000);
                }
            } else if (freq === 'weekly') {
                // 默认周一
                while (next.getDay() !== 1) {
                    next = new Date(next.getTime() + 24 * 3600 * 1000);
                }
            }
            return this._fmtYmdHms(next);
        },

        /** 格式化 Date → "yyyy-MM-dd HH:mm:ss" */
        _fmtYmdHms(d) {
            const pad = (n) => String(n).padStart(2, '0');
            return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate())
                + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
        },

        /**
         * 后台同步到 HTTP API（不阻塞 UI，失败不影响本地数据）
         */
        _syncToBackend(taskData, editingId) {
            // 先取 9528 的 session token 再推（避免 /api/* 401；9529 无需 token）
            getExcelToken().then(() => {
                const ctrl = new AbortController();
                const t = setTimeout(() => ctrl.abort(), 3000);
                const url = editingId ? API_BASE + '/tasks/' + editingId : API_BASE + '/tasks';
                const method = editingId ? 'PUT' : 'POST';
                fetch(excelApiUrl(url), {
                    method,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(taskData),
                    signal: ctrl.signal
                }).then(() => clearTimeout(t)).catch(() => clearTimeout(t));
            });

            // 双通道：也尽力推送到独立服务 SAR 9529（fire-and-forget，失败静默）
            // TaskSchedulerBridge 在后端也会做同样推送；这里是“用户在侧栏保存后立刻想在独立页看到”的立即同步
            try {
                const sarUrl = (App.Endpoints ? App.Endpoints.standaloneAgent.tasks : 'http://localhost:9529/tasks');
                const sarBody = Object.assign({}, taskData, { Id: editingId || taskData.Id });
                const c2 = new AbortController();
                const t2 = setTimeout(() => c2.abort(), 2500);
                fetch(sarUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(sarBody),
                    signal: c2.signal
                }).then(() => clearTimeout(t2)).catch(() => clearTimeout(t2));
            } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
        },

        /**
         * 打开新建表单（内联面板）
         */
        openCreateForm() {
            this._editingId = null;
            // 确保面板打开
            if (typeof window.toggleTaskPlanner === 'function') {
                const tpPanel = document.getElementById('task-planner-panel');
                if (!tpPanel || !tpPanel.classList.contains('active')) {
                    window.toggleTaskPlanner();
                }
            }
            // 填充内联表单
            document.getElementById('tp-form-title').textContent = '新建定时任务';
            document.getElementById('tp-name').value = '';
            document.getElementById('tp-frequency').value = 'daily';
            document.getElementById('tp-time').value = '09:00';
            document.getElementById('tp-command').value = '';
            const tgtNew = document.getElementById('tp-target-file');
            if (tgtNew) tgtNew.value = '';
            // 默认目标文件：宿主（当前 Excel 打开的工作簿）
            (async () => {
                try {
                    let host = null;
                    // COM
                    try {
                        if (window.ExcelBridge && typeof window.ExcelBridge.GetHostWorkbook === 'function') {
                            const raw = window.ExcelBridge.GetHostWorkbook();
                            if (raw) host = (typeof raw === 'string') ? JSON.parse(raw) : raw;
                        }
                    } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
                    // HTTP fallback
                    if (!host && window.fetch) {
                        try {
                            await getExcelToken();
                            const res = await fetch(excelApiUrl((App.Endpoints ? App.Endpoints.excelVstoHttp.hostWorkbook : 'http://localhost:9528/api/host/workbook')), { method: 'GET' });
                            if (res.ok) host = await res.json();
                        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
                    }
                    if (host && host.hasHost && host.path) {
                        const el = document.getElementById('tp-target-file');
                        if (el && !el.value) el.value = host.path;
                    }
                } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
            })();
            document.getElementById('tp-time-group').style.display = '';
            // 任务级模型选项：重置选项 + 自动刷新（AUTO / 本地 Ollama / 云端预设）
            try {
                const modelSel = document.getElementById('tp-task-model');
                if (modelSel) { modelSel.value = ''; TaskPlanner.refreshTaskModelSelect('tp-task-model'); }
            } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
            document.getElementById('tp-inline-form').style.display = 'block';
            // 聚焦到名称输入框
            setTimeout(() => {
                const nameInput = document.getElementById('tp-name');
                if (nameInput) nameInput.focus();
            }, 100);
        },

        /**
         * 打开编辑表单（内联面板）
         */
        _edit(id) {
            const task = this._tasks.find(t => t.Id === id);
            if (!task) return;
            this._editingId = id;
            // 确保面板打开
            if (typeof window.toggleTaskPlanner === 'function') {
                const tpPanel = document.getElementById('task-planner-panel');
                if (!tpPanel || !tpPanel.classList.contains('active')) {
                    window.toggleTaskPlanner();
                }
            }
            document.getElementById('tp-form-title').textContent = '编辑定时任务';
            document.getElementById('tp-name').value = task.Name || '';
            document.getElementById('tp-frequency').value = task.Frequency || 'daily';
            document.getElementById('tp-time').value = task.Time || '09:00';
            document.getElementById('tp-command').value = task.Command || '';
            const tgtEdit = document.getElementById('tp-target-file');
            if (tgtEdit) tgtEdit.value = task.TargetFile || task.target_file || '';
            document.getElementById('tp-time-group').style.display = task.Frequency === 'hourly' ? 'none' : '';
            // 任务级模型：刷新选项并回填当前任务的值
            try {
                const modelSel = document.getElementById('tp-task-model');
                if (modelSel) {
                    const cur = task.Model || task.model || '';
                    // 先占位：刷新选项后再选回 cur
                    TaskPlanner.refreshTaskModelSelect('tp-task-model', null, cur);
                }
            } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
            document.getElementById('tp-inline-form').style.display = 'block';
        },

        /**
         * 提交内联表单
         */
        async submitInlineForm() {
            const name = document.getElementById('tp-name').value.trim();
            const frequency = document.getElementById('tp-frequency').value;
            const time = document.getElementById('tp-time').value || '09:00';
            const command = document.getElementById('tp-command').value.trim();
            const tgtEl = document.getElementById('tp-target-file');
            const targetFile = tgtEl ? (tgtEl.value || '').trim() : '';
            // 任务级模型：不选=空串(AUTO)；选了=具体模型名（带 provider 前缀？保持模型名即可，由独立服务 / 执行层判断）
            const modelEl = document.getElementById('tp-task-model');
            const modelVal = modelEl ? (modelEl.value || '').trim() : '';

            if (!name) { this._toast('请输入任务名称', 'error'); return; }
            if (!command) { this._toast('请输入执行指令', 'error'); return; }

            const saveBtn = document.getElementById('tp-save-btn');
            if (saveBtn) {
                saveBtn.disabled = true;
                saveBtn.textContent = '保存中...';
                saveBtn.style.opacity = '0.7';
            }

            const safetyTimer = setTimeout(() => {
                if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = '保存'; saveBtn.style.opacity = '1'; }
            }, 10000);

            const taskData = { Name: name, Frequency: frequency, Time: time, Command: command, Enabled: true, TargetFile: targetFile, Model: modelVal };

            try {
                let success = false;
                let errorMsg = '';
                // COM 分支拿到的服务端最终数据（含 NextRun/Id），稍后合并入本地
                let serverData = null;

                // 优先通过 COM Bridge
                try {
                    if (window.taskScheduler) {
                        const json = JSON.stringify(taskData);
                        let result = null;
                        if (this._editingId) {
                            result = window.taskScheduler.UpdateTask(this._editingId, json);
                        } else {
                            result = window.taskScheduler.CreateTask(json);
                        }
                        if (result) {
                            const parsed = JSON.parse(result);
                            success = parsed.success === true;
                            if (success) serverData = parsed.data || null;
                            if (!success) errorMsg = parsed.error || parsed.message || '';
                        } else {
                            errorMsg = 'COM Bridge 返回空结果';
                        }
                    } else {
                        errorMsg = 'COM Bridge 不可用';
                    }
                } catch (e) {
                    errorMsg = 'COM Bridge 调用失败: ' + e.message;
                }

                let persistedId = this._editingId;

                if (success && serverData && serverData.Id) {
                    // COM 成功 → 用服务端回传的完整对象（含 NextRun）合并到本地列表
                    persistedId = serverData.Id;
                    const idx = this._tasks.findIndex(t => String(t.Id) === String(serverData.Id));
                    const merged = idx >= 0 ? { ...this._tasks[idx], ...serverData } : { ...serverData };
                    // 兜底：后端万一没返回 NextRun，前端重算（保证“保存后马上看到下次执行”）
                    if (!merged.NextRun) merged.NextRun = this.calcNextRun(merged);
                    if (idx >= 0) this._tasks[idx] = merged; else this._tasks.push(merged);
                    this._persistLocal();
                } else if (!success) {
                    // localStorage 回退
                    const localTask = { ...taskData };
                    if (this._editingId) {
                        localTask.Id = this._editingId;
                        const idx = this._tasks.findIndex(t => t.Id === this._editingId);
                        // 前端重算 NextRun → 保存后立刻看到新的下次执行时间
                        const existing = idx >= 0 ? { ...this._tasks[idx], ...localTask } : localTask;
                        existing.NextRun = this.calcNextRun(existing);
                        if (idx >= 0) this._tasks[idx] = existing; else this._tasks.push(existing);
                    } else {
                        localTask.Id = this._genId();
                        localTask.CreatedAt = new Date().toISOString();
                        localTask.NextRun = this.calcNextRun(localTask);
                        this._tasks.push(localTask);
                    }
                    persistedId = localTask.Id;
                    this._persistLocal();
                    // 后台重试 COM Bridge 同步，确保定时器启动
                    if (window.taskScheduler) {
                        try {
                            var retryResult = null;
                            if (this._editingId) {
                                retryResult = window.taskScheduler.UpdateTask(this._editingId, JSON.stringify(taskData));
                            }
                            if (!retryResult) {
                                retryResult = window.taskScheduler.CreateTask(JSON.stringify(taskData));
                            }
                            if (retryResult) {
                                const parsed = JSON.parse(retryResult);
                                if (parsed.success && parsed.data) {
                                    const sId = parsed.data.Id;
                                    const idx = this._tasks.findIndex(t => String(t.Id) === String(sId || localTask.Id));
                                    if (idx >= 0) {
                                        this._tasks[idx] = { ...this._tasks[idx], ...parsed.data };
                                        if (!this._tasks[idx].NextRun) this._tasks[idx].NextRun = this.calcNextRun(this._tasks[idx]);
                                        this._persistLocal();
                                    }
                                }
                            }
                        } catch (e) { /* 后台同步失败不影响本地数据 */ }
                    }
                    success = true;
                }

                clearTimeout(safetyTimer);

                if (success) {
                    this._toast(this._editingId ? '任务已更新' : '任务已创建', 'success');
                    // 若已写入 this._tasks 则用 persistedId，否则仍按原 taskData 推送
                    const taskForSync = (persistedId ? (this._tasks.find(t => String(t.Id) === String(persistedId)) || null) : null) || taskData;
                    this._syncToBackend(taskForSync, this._editingId);
                    // 隐藏内联表单
                    if (typeof window.hideInlineForm === 'function') {
                        window.hideInlineForm();
                    } else {
                        document.getElementById('tp-inline-form').style.display = 'none';
                    }
                    this._editingId = null;
                    this._renderTasks();
                    this._renderStats();
                } else {
                    if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = '保存'; saveBtn.style.opacity = '1'; }
                    this._toast('操作失败: ' + errorMsg, 'error');
                }
            } catch (e) {
                if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = '保存'; saveBtn.style.opacity = '1'; }
                this._toast('保存异常: ' + e.message, 'error');
            }
        },

        /**
         * 刷新任务列表
         */
        refresh() {
            this.loadTasks();
        },

        /**
         * 渲染统计信息
         */
        _renderStats() {
            const statsEl = document.getElementById('tp-stats');
            if (!statsEl) return;
            const total = this._tasks.length;
            const running = this._tasks.filter(t => t.Enabled !== false).length;
            const paused = total - running;
            const ti = (k) => { try { return (window.App && App.i18n && App.i18n.t) ? App.i18n.t(k) : k; } catch (_) { return k; } };
            statsEl.innerHTML = `
                <div style="flex:1;background:#fff;border-radius:8px;padding:10px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.06);">
                    <div style="font-size:22px;font-weight:700;color:#d32f2f;">${total}</div>
                    <div style="font-size:11px;color:#888;margin-top:2px;">${ti('tp.all')}</div>
                </div>
                <div style="flex:1;background:#fff;border-radius:8px;padding:10px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.06);">
                    <div style="font-size:22px;font-weight:700;color:#4caf50;">${running}</div>
                    <div style="font-size:11px;color:#888;margin-top:2px;">${ti('tp.running')}</div>
                </div>
                <div style="flex:1;background:#fff;border-radius:8px;padding:10px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.06);">
                    <div style="font-size:22px;font-weight:700;color:#ff9800;">${paused}</div>
                    <div style="font-size:11px;color:#888;margin-top:2px;">${ti('tp.paused')}</div>
                </div>
            `;
        },

        /**
         * 渲染任务列表
         */
        _renderTasks() {
            const listEl = document.getElementById(this._listElId) || document.getElementById('tasks-list');
            if (!listEl) return;

            // 同时渲染统计
            this._renderStats();

            if (this._tasks.length === 0) {
                listEl.innerHTML = '<div style="text-align:center;padding:40px 20px;color:#999;font-size:13px;">暂无定时任务<br><span style="font-size:12px;color:#bbb;">点击"新建任务"按钮创建第一个定时任务</span></div>';
                return;
            }

            listEl.innerHTML = this._tasks.map(t => {
                const ti = (k) => { try { return (window.App && App.i18n && App.i18n.t) ? App.i18n.t(k) : k; } catch (_) { return k; } };
                const statusDot = t.Enabled ? 'running' : 'paused';
                const statusText = t.Enabled ? ti('tp.running') : ti('tp.paused');
                const freqText = {
                    once: ti('tp.freqOnce'),
                    hourly: ti('tp.freqHourly'),
                    daily: ti('tp.freqDaily'),
                    weekdays: ti('tp.freqWeekdays'),
                    weekly: ti('tp.freqWeekly')
                }[t.Frequency] || t.Frequency;
                const nextRun = t.NextRun || '--';
                const lastRun = t.LastRun || '从未执行';

                const taskModelRaw = t.Model || t.model || '';
                const taskModel = this.parseTaskModelValue(taskModelRaw);
                let modelBadge = '';
                if (taskModel && taskModel.model) {
                    const providerLabels = { openai:'OpenAI', azure:'Azure', moonshot:'Kimi', qianwen:'千问', zhipu:'智谱', ollama:'Ollama', deepseek:'DeepSeek', anthropic:'Anthropic' };
                    const pLabel = taskModel.provider ? (providerLabels[taskModel.provider] || taskModel.provider) : '';
                    const isVision = /(vision|vl|vt|llava|minicpm|pixtral|gemini|internvl)/i.test(taskModel.model);
                    const isReason = /(r1|reasoning|deepseek-reasoner|o1|o3)/i.test(taskModel.model);
                    const badgeColor = taskModel.provider === 'ollama' ? '#3c823d' : (isReason ? '#7c3aed' : (isVision ? '#1976d2' : '#d32f2f'));
                    const prefix = (taskModel.provider === 'ollama' ? '🖥️ ' : '') + (isVision ? '👁️ ' : '') + (isReason ? '🧠 ' : '');
                    const inner = prefix + (pLabel ? pLabel + ' · ' : '') + taskModel.model;
                    modelBadge = `<div style="margin-top:6px;display:inline-block;"><span style="display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;color:#fff;background:${badgeColor};vertical-align:middle;">${this._escapeHtml(inner)}</span></div>`;
                }
                return `<div style="background:#fff;border-radius:8px;padding:12px 14px;margin-bottom:8px;box-shadow:0 1px 4px rgba(0,0,0,0.08);">
                    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
                        <div style="font-weight:600;font-size:14px;display:flex;align-items:center;gap:6px;">
                            <span style="width:8px;height:8px;border-radius:50%;display:inline-block;background:${t.Enabled?'#4caf50':'#ff9800'};"></span>
                            ${this._escapeHtml(t.Name || '未命名任务')}
                        </div>
                        <div style="display:flex;gap:4px;">
                            <button onclick="window.TaskPlanner._execute('${t.Id}')" style="padding:3px 8px;border:none;border-radius:4px;background:#4caf50;color:#fff;cursor:pointer;font-size:11px;" title="立即执行">▶</button>
                            <button onclick="window.TaskPlanner._toggle('${t.Id}')" style="padding:3px 8px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;font-size:11px;">${t.Enabled?'⏸':'▶'}</button>
                            <button onclick="window.TaskPlanner._edit('${t.Id}')" style="padding:3px 8px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;font-size:11px;">✏️</button>
                            <button onclick="window.TaskPlanner._delete('${t.Id}')" style="padding:3px 8px;border:1px solid #f44336;border-radius:4px;background:#fff;color:#f44336;cursor:pointer;font-size:11px;">🗑</button>
                        </div>
                    </div>
                    <div style="font-size:12px;color:#666;line-height:1.7;">
                        <span style="color:#999;">频率：</span><strong>${freqText}</strong>
                        <span style="color:#999;margin-left:12px;">时间：</span><strong>${t.Time || '--'}</strong>
                        <span style="color:#999;margin-left:12px;">状态：</span><strong>${statusText}</strong>
                    </div>
                    <div style="font-size:12px;color:#666;">
                        <span style="color:#999;">下次执行：</span><strong>${nextRun}</strong>
                        <span style="color:#999;margin-left:12px;">上次执行：</span><strong>${lastRun}</strong>
                    </div>
                    ${t.Command ? `<div style="margin-top:6px;font-size:12px;color:#888;background:#f9f9f9;padding:4px 8px;border-radius:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${this._escapeHtml(t.Command)}</div>` : ''}
                    ${modelBadge}
                    ${t.TargetFile ? `<div style="margin-top:6px;font-size:12px;color:#1a73e8;background:#e8f0fe;padding:4px 8px;border-radius:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${this._escapeHtml(t.TargetFile)}">📄 目标：${this._escapeHtml(t.TargetFile)}</div>` : ''}
                </div>`;
            }).join('');
        },

        /**
         * 执行任务
         */
        async _execute(id) {
            const task = this._tasks.find(t => t.Id === id);
            if (!task) { this._toast('任务不存在', 'error'); return; }

            try {
                let success = false;
                // 1. 尝试 COM Bridge 执行
                try {
                    if (window.taskScheduler?.ExecuteTask) {
                        const result = window.taskScheduler.ExecuteTask(id);
                        if (result) {
                            const parsed = JSON.parse(result);
                            success = parsed.success || parsed.data;
                        }
                    }
                } catch (e) { /* fallback */ }

                // 2. 尝试 HTTP API 执行
                if (!success) {
                    try {
                        await getExcelToken();
                        const res = await fetch(excelApiUrl(API_BASE + '/tasks/' + id + '/execute'), { method: 'POST' });
                        const data = await res.json();
                        success = data.success || data.data;
                    } catch (e) { /* fallback */ }
                }

                // 3. 本地实际执行：将指令填入聊天输入框并发送
                const command = task.Command || '';
                if (command) {
                    // 设置绕过守卫标志（用户主动点击执行，已默认授权）
                    window._tpBypassGuard = true;

                    const input = document.getElementById('message-input');
                    if (input) {
                        input.value = command;
                        // 触发 input 事件，让绑定层感知值变化
                        const evt = new Event('input', { bubbles: true });
                        input.dispatchEvent(evt);
                        // 自动发送
                        if (typeof sendMessage === 'function') {
                            setTimeout(() => {
                                sendMessage();
                                // 30 秒后清除绕过标志
                                setTimeout(() => {
                                    window._tpBypassGuard = false;
                                }, 30000);
                            }, 300);
                        } else {
                            window._tpBypassGuard = false;
                        }
                    } else {
                        window._tpBypassGuard = false;
                    }
                }

                // 4. 更新执行记录
                task.LastRun = new Date().toLocaleString('zh-CN');
                this._persistLocal();
                this._renderTasks();
                success = true;

                this._toast('任务已触发执行', 'success');
            } catch (e) {
                this._toast('执行异常: ' + e.message, 'error');
            }
        },

        /**
         * 切换暂停/恢复
         */
        async _toggle(id) {
            const task = this._tasks.find(t => t.Id === id);
            if (!task) return;

            try {
                let success = false;
                try {
                    if (window.taskScheduler?.UpdateTask) {
                        const result = window.taskScheduler.UpdateTask(id, JSON.stringify({ Enabled: !task.Enabled }));
                        if (result) {
                            const parsed = JSON.parse(result);
                            success = parsed.success || parsed.data;
                        }
                    }
                } catch (e) { /* fallback */ }

                if (!success) {
                    await getExcelToken();
                    const res = await fetch(excelApiUrl(API_BASE + '/tasks/' + id), {
                        method: 'PUT',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ Enabled: !task.Enabled })
                    });
                    const data = await res.json();
                    success = data.success || data.data;
                }

                if (!success) {
                    // 最终回退：localStorage
                    task.Enabled = !task.Enabled;
                    this._persistLocal();
                    success = true;
                }

                if (success) {
                    this._toast(task.Enabled ? '任务已恢复' : '任务已暂停', 'success');
                    this._renderTasks();
                } else {
                    this._toast('操作失败', 'error');
                }
            } catch (e) {
                this._toast('网络错误: ' + e.message, 'error');
            }
        },

        /**
         * 删除任务
         */
        async _delete(id) {
            if (!confirm('确定要删除此任务吗？')) return;

            try {
                let success = false;
                let errorMsg = '';

                try {
                    if (window.taskScheduler?.DeleteTask) {
                        const result = window.taskScheduler.DeleteTask(id);
                        if (result) {
                            const parsed = JSON.parse(result);
                            success = parsed.success === true;
                            if (!success) errorMsg = parsed.error || parsed.message || '';
                        } else {
                            errorMsg = 'COM Bridge 返回空结果';
                        }
                    } else {
                        errorMsg = 'COM Bridge 不可用';
                    }
                } catch (e) {
                    errorMsg = 'COM Bridge 调用失败: ' + e.message;
                }

                if (!success) {
                    try {
                        await getExcelToken();
                        const res = await fetch(excelApiUrl(API_BASE + '/tasks/' + id), { method: 'DELETE' });
                        const data = await res.json();
                        success = data.success === true;
                        if (!success) errorMsg = data.error || data.message || '';
                    } catch (e2) {
                        errorMsg = 'HTTP API 失败: ' + e2.message;
                    }
                }

                if (!success) {
                    // 最终回退：localStorage
                    success = true;
                }

                if (success) {
                    // 无论哪种方式成功，都从本地移除（防止后端删了本地还在）
                    this._tasks = this._tasks.filter(t => t.Id !== id);
                    this._persistLocal();
                    this._toast('任务已删除', 'success');
                    this._renderTasks();
                } else {
                    this._toast('删除失败: ' + errorMsg, 'error');
                }
            } catch (e) {
                this._toast('删除异常: ' + e.message, 'error');
            }
        },

        /**
         * Toast 提示
         */
        _toast(msg, type) {
            const toast = document.createElement('div');
            toast.id = 'tp-toast-' + Date.now();
            toast.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);color:#fff;padding:14px 28px;border-radius:8px;font-size:15px;font-weight:600;z-index:100000;box-shadow:0 4px 16px rgba(0,0,0,0.3);';
            if (type === 'success') toast.style.background = '#4caf50';
            else if (type === 'error') toast.style.background = '#f44336';
            else toast.style.background = '#333';
            toast.textContent = msg;
            document.body.appendChild(toast);
            setTimeout(() => toast.remove(), 3000);
        },

        /**
         * HTML 转义
         */
        _escapeHtml(str) {
            if (!str) return '';
            return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
        },

        /**
         * 按 Name+Frequency+Time 指纹去重，保留首次出现的任务
         */
        _deduplicate() {
            try {
                const raw = localStorage.getItem(this._STORAGE_KEY);
                if (!raw) return;
                const tasks = JSON.parse(raw);
                if (!Array.isArray(tasks) || tasks.length < 2) return;
                const seen = new Set();
                const deduped = [];
                let removed = 0;
                for (const t of tasks) {
                    const key = (t.Name||'') + '|' + (t.Frequency||'') + '|' + (t.Time||'');
                    if (seen.has(key)) { removed++; continue; }
                    seen.add(key);
                    deduped.push(t);
                }
                if (removed > 0) {
                    localStorage.setItem(this._STORAGE_KEY, JSON.stringify(deduped));
                    console.log('[TaskPlanner] 去重完成，移除 ' + removed + ' 个重复任务');
                }
            } catch (e) {
                console.warn('[TaskPlanner] 去重失败:', e.message);
            }
        },

        /**
         * 任务级「执行模型」下拉：按聊天栏相同的分组方式渲染（AUTO / Ollama 本地 / 云端各服务商）。
         *   - selectId：<select> 的 DOM id（内嵌面板=tp-task-model；独立页=create-form-model / edit-form-model）
         *   - btnElOpt：按钮元素，点刷新时禁用
         *   - restoreValue：刷新后需要恢复选中的值（编辑表单时用）
         * Provider 名/默认模型列表：保持和 PROVIDER_CONFIGS 中主流服务商一致即可（llm_config 只认 provider + model 字段）
         */
        _TP_PROVIDER_DEFAULTS: [
            { key: 'openai',    label: 'OpenAI',        models: ['gpt-5.6-sol','gpt-5.6-terra','gpt-5.6-luna','gpt-5.5','gpt-5.4','gpt-5.4-pro','gpt-5.4-mini','gpt-5.4-nano','gpt-5.3-codex','gpt-4.1','gpt-4.1-mini','gpt-4o','gpt-4o-mini','gpt-6-astra','gpt-chat-latest'] },
            { key: 'deepseek',  label: 'DeepSeek',      models: ['deepseek-v4-pro','deepseek-v4-flash'] },
            { key: 'volcengine', label: '火山引擎',     models: ['doubao-seed-2-1-pro-260628','doubao-seed-2-0-pro-260215','doubao-seed-2-0-lite-260215','doubao-seed-2-0-mini-260215','doubao-seed-2-0-code-preview-260215','doubao-seed-1-8-251228','doubao-1-5-pro-32k-character-250228'] },
            { key: 'moonshot',  label: 'Kimi 月之暗面', models: ['kimi-k3','kimi-k2.7-code','kimi-k2.7-code-highspeed','kimi-k2.6'] },
            { key: 'qianwen',   label: '千问通义',      models: ['qwen3.8-max','qwen3.8-flash','qwen3.8-max-0902','qwen3.7-max','qwen3.7-plus','qwen3.6-flash','qwen3-max','qwen3-coder-plus'] },
            { key: 'zhipu',     label: '智谱 GLM',      models: ['glm-5.3','glm-5.3-flash','glm-5.2','glm-5.1','glm-5.1-highspeed','glm-5','glm-5-turbo','glm-5v-turbo','glm-4.7','glm-4.6v'] },
            { key: 'anthropic', label: 'Anthropic',     models: ['claude-3-5-sonnet-latest','claude-3-5-sonnet-20241022','claude-3-5-haiku-latest','claude-3-opus-latest'] },
            { key: 'azure',     label: 'Azure OpenAI',  models: ['gpt-5.6-sol','gpt-5.6-terra','gpt-5.6-luna','gpt-5.4','gpt-5.4-mini','gpt-4.1','gpt-4.1-mini','gpt-4o','gpt-4o-mini','gpt-oss-120b','gpt-oss-20b'] },
            // ollama 单独 detect，默认兜底列表（仅独立页 file:// 无 PROVIDER_CONFIGS 时使用）
            { key: 'ollama',    label: 'Ollama 本地',   models: ['llama3.1','llama3.2','llama3.3','qwen2.5:7b','qwen2.5:32b','qwen2.5:72b','qwen3:4b','qwen3:8b','qwen3:30b','deepseek-r1:8b','gemma3:4b','gemma3:12b','qwen2.5-vl'] }
        ],

        /**
         * 构建任务级模型下拉的分组列表，来源与聊天栏保持一致：
         *   - 内嵌面板 / Excel 运行中（存在全局 PROVIDER_CONFIGS）：
         *     复用 PROVIDER_CONFIGS 并施加与聊天栏相同的可用性过滤（云端需已配置 API Key；
         *     Ollama 需检测到本地模型或配置了 localhost URL），未配置的服务商不出现——与聊天区一致；
         *   - 独立页 file://（Excel 关闭，无 PROVIDER_CONFIGS）：回退 _TP_PROVIDER_DEFAULTS。
         * 返回 [{ key, label, models: [string] }]，Ollama 在前。
         */
        _listTaskModelGroups() {
            // 1) 内嵌面板 —— 直接复用聊天区真实的模型分组（DOM → groups 结构）：
            //    提供商是否可用、模型是否真实存在（Ollama 用 /api/tags 探测），都由聊天区那套判断好了。
            //    从这里抽 = 跟聊天区选项 100% 一致，不再自己造一份"看似能用实际不可达"的假大列表。
            const chatSel = document.getElementById('chat-model-select');
            if (chatSel && chatSel.options && chatSel.options.length > 0) {
                const groups = [];
                let cur = null;
                for (let i = 0; i < chatSel.options.length; i++) {
                    const opt = chatSel.options[i];
                    // AUTO 顶部不进入分组，独立在 append 时前面加；其它所有分组/模型按顺序抄
                    const value = String(opt.value || '');
                    const text = String(opt.textContent || '').replace(/^[\s─]+/, '').replace(/[\s─]+$/, '');
                    // 分组标签：value 为空 + disabled = true，文本形如 ─── OpenAI ───
                    if (!value && opt.disabled) {
                        const labelOrig = text;
                        if (labelOrig) {
                            // 从 label 反查 provider key（聊天区用 getProviderName，这里用 ChatLog/App.getProviderName 反查
                            let key = null;
                            if (typeof App !== 'undefined' && App && typeof App.getProviderName === 'function') {
                                const keys = Object.keys(typeof PROVIDER_CONFIGS !== 'undefined' ? PROVIDER_CONFIGS : {});
                                for (const k of keys) {
                                    if (App.getProviderName(k) === labelOrig || App.getProviderName(k).replace(/^\s+|\s+$/g,'') === labelOrig) { key = k; break; }
                                }
                                if (!key) {
                                    // 常见兜底匹配：label 可能带"本地/云端"前缀修饰，从选项里反推 provider（同组内非 disabled 选项的 data.provider）
                                    for (let j = i + 1; j < chatSel.options.length; j++) {
                                        const o2 = chatSel.options[j];
                                        if (String(o2.value || '') && !o2.disabled && o2.dataset && o2.dataset.provider) { key = o2.dataset.provider; break; }
                                        if (String(o2.value || '') === '' && o2.disabled) break;   // 下一组
                                    }
                                }
                            }
                            cur = { key: key || ('group_' + i), label: labelOrig, models: [] };
                            groups.push(cur);
                        }
                        continue;
                    }
                    if (!value || value === 'AUTO' || value === 'auto') continue;
                    const providerKey = (opt.dataset && opt.dataset.provider) ? String(opt.dataset.provider)
                        : (cur && cur.key ? cur.key : '');
                    if (!cur || (cur.key !== providerKey && providerKey)) {
                        let label = providerKey;
                        if (typeof App !== 'undefined' && App && typeof App.getProviderName === 'function') {
                            try { label = App.getProviderName(providerKey); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
                        }
                        cur = { key: providerKey, label: label, models: [] };
                        groups.push(cur);
                    }
                    if (cur) cur.models.push(value);
                }
                if (groups.length) return groups;
            }

            // 2) 无聊天区 DOM（独立页/初始化未就绪）→ 回退 PROVIDER_CONFIGS 或 _TP_PROVIDER_DEFAULTS
            const hasGlobal = (typeof PROVIDER_CONFIGS !== 'undefined') && PROVIDER_CONFIGS;
            if (hasGlobal) {
                const cfg = (typeof config !== 'undefined' && config) ? config : {};
                const avail = (typeof App !== 'undefined' && App && typeof App.getAvailableProviders === 'function')
                    ? App.getAvailableProviders()
                    : Object.keys(PROVIDER_CONFIGS).filter(k => {
                        const pc = PROVIDER_CONFIGS[k];
                        if (!pc) return false;
                        if (k === 'ollama') {
                            return (detectedModels && detectedModels.length) ||
                                String((cfg.apiUrl && String(cfg.apiUrl)) || pc.apiUrl || '').includes('localhost');
                        }
                        if (pc.needsKey) { return !!(cfg.providers && cfg.providers[k] && String(cfg.providers[k].apiKey).trim()); }
                        return true;
                      });
                const providerLabel = (key) => {
                    if (typeof App !== 'undefined' && App && typeof App.getProviderName === 'function') {
                        try { return App.getProviderName(key); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
                    }
                    const names = { ollama:'Ollama 本地',openai:'OpenAI',azure:'Azure OpenAI',wenxin:'文心一言',qianwen:'通义千问',zhipu:'智谱 AI',xunfei:'讯飞星火',moonshot:'Moonshot',custom:'自定义',deepseek:'DeepSeek',anthropic:'Anthropic' };
                    return names[key] || key;
                };
                const normalize = (m) => (typeof m === 'string' ? m : (m && m.name ? m.name : ''));
                const groups = [];
                const pushGroup = (key) => {
                    if (!avail.includes(key)) return;
                    const pc = PROVIDER_CONFIGS[key];
                    if (!pc) return;
                    let models = [];
                    if (key === 'ollama' && detectedModels && detectedModels.length > 0) {
                        models = detectedModels.map(m => (m.name || m.model || '').toString()).filter(Boolean);
                    } else if (key === 'ollama' && !(detectedModels && detectedModels.length)) {
                        // Ollama 没探测到任何模型 → 不把预置模型当成"能用"展示
                        models = [];
                    } else {
                        models = (pc.models || []).map(normalize).filter(Boolean);
                    }
                    if (models.length) groups.push({ key, label: providerLabel(key), models });
                };
                if (PROVIDER_CONFIGS['ollama']) pushGroup('ollama');
                Object.keys(PROVIDER_CONFIGS).forEach(k => { if (k !== 'ollama' && k !== 'custom') pushGroup(k); });
                return groups;
            }
            return this._TP_PROVIDER_DEFAULTS.map(p => ({ key: p.key, label: p.label, models: p.models.slice() }));
        },

        async refreshTaskModelSelect(selectId, btnElOpt, restoreValue) {
            const sel = document.getElementById(selectId);
            if (!sel) return;
            if (btnElOpt) { btnElOpt.disabled = true; const t = btnElOpt.textContent; btnElOpt.textContent = '刷新中…'; try { btnElOpt._origText = t; } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); } }
            try {
                // 刷新任务模型前先刷新聊天区的模型选择器（这才是"真正检测到的源"）；
                // 因为任务表单直接复用聊天区 DOM → 必须让它先重算一遍，否则刷新是假的。
                try {
                    if (typeof App !== 'undefined' && App && typeof App.refreshModelsList === 'function') {
                        await Promise.resolve(App.refreshModelsList()).catch(() => null);
                    }
                    if (typeof App !== 'undefined' && App && typeof App.updateChatModelSelector === 'function') {
                        App.updateChatModelSelector();
                    }
                } catch(e) { /* 初始化阶段函数未就绪，忽略 */ }

                // 1) 先尝试探测本地 Ollama（11434）拉真实模型；失败回退
                let ollamaModels = [];
                try {
                    const c = new AbortController(); const tt = setTimeout(()=>c.abort(), 2000);
                    const r = await fetch('http://localhost:11434/api/tags', { method:'GET', signal: c.signal });
                    clearTimeout(tt);
                    if (r.ok) {
                        const jr = await r.json();
                        if (jr && Array.isArray(jr.models)) {
                            ollamaModels = jr.models.map(m => {
                                let n = (m.name || m.model || '').toString();
                                if (n && !n.includes(':')) n = n + ':latest';
                                return n;
                            }).filter(Boolean);
                        }
                    }
                } catch(e) { /* Ollama 不在线就走分组列表 */ }

                // 2) 分组列表来源：与聊天栏一致的 PROVIDER_CONFIGS + 可用性过滤；独立页回退默认
                let groups = this._listTaskModelGroups();
                // 若实时探测到 Ollama 模型，覆盖 ollama 分组（与聊天栏“先检测真实模型”一致）
                if (ollamaModels.length) {
                    const gi = groups.findIndex(g => g.key === 'ollama');
                    if (gi >= 0) groups[gi].models = ollamaModels;
                    else groups.unshift({ key: 'ollama', label: 'Ollama 本地', models: ollamaModels });
                }

                // 3) 清掉旧选项，重建分组：AUTO 开头 → Ollama 本地 → 云端
                sel.innerHTML = '';
                const autoOpt = document.createElement('option');
                autoOpt.value = '';
                autoOpt.textContent = '🤖 AUTO（智能自动选择）';
                sel.appendChild(autoOpt);

                const appendGroup = (groupLabel, models, providerKey) => {
                    if (!models || !models.length) return;
                    const grp = document.createElement('optgroup');
                    grp.label = '── ' + groupLabel + (providerKey === 'ollama' ? '（' + models.length + '）' : '') + ' ──';
                    // 用不可选的 <option> 样式分组标签 + 选项（optgroup 在部分 WebView2 中渲染有差异，保险起见两种都加）
                    const labelOpt = document.createElement('option');
                    labelOpt.value = '';
                    labelOpt.disabled = true;
                    labelOpt.style.cssText = 'font-weight:700;background:#f5f5f5;color:#333;';
                    labelOpt.textContent = (providerKey === 'ollama' ? '🖥️  ' : '☁️  ') + groupLabel + (providerKey === 'ollama' ? '  ·  本地检测到 ' + models.length + ' 个' : '');
                    sel.appendChild(labelOpt);
                    const unique = []; const seen = new Set();
                    for (const m of models) { if (!seen.has(m)) { seen.add(m); unique.push(m); } }
                    for (const m of unique) {
                        const o = document.createElement('option');
                        o.value = providerKey + '::' + m;   // provider::model，后端拆开
                        o.dataset.provider = providerKey;
                        o.dataset.model = m;
                        // 图标标识
                        let prefix = '  · ';
                        if (providerKey === 'ollama') prefix = '   🖥️ ';
                        const isVision = /(vision|vl|vt|llava|minicpm|pixtral|gemini|internvl)/i.test(m);
                        const isReason = /(r1|reasoning|deepseek-reasoner|o1|o3)/i.test(m);
                        if (isVision) prefix += '👁️ ';
                        if (isReason) prefix += '🧠 ';
                        o.textContent = prefix + m;
                        if (isVision) o.style.color = '#4a90d9';
                        if (isReason) o.style.color = '#7c3aed';
                        sel.appendChild(o);
                    }
                };

                // 4) Ollama 本地优先，然后云端（保持与聊天栏顺序一致）
                groups.forEach(g => {
                    appendGroup(g.label, g.models, g.key);
                });

                // 5) 回填 restoreValue
                const rv = restoreValue || '';
                if (rv) {
                    const exists = Array.from(sel.options).some(o => o.value === rv);
                    if (exists) { sel.value = rv; }
                    else {
                        // 兼容老格式：裸模型名 / {provider,model} JSON / "model" 字符串
                        try {
                            let guess = rv;
                            if (/^[\[{]/.test(rv)) {
                                const j = JSON.parse(rv);
                                guess = (j.provider || '') + '::' + (j.model || '');
                            }
                            if (!guess.includes('::')) { guess = '::' + guess; }
                            const exists2 = Array.from(sel.options).some(o => o.value === guess);
                            if (exists2) sel.value = guess;
                        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
                    }
                }
            } catch (e) {
                console.warn('[TaskPlanner] refreshTaskModelSelect failed:', e.message);
            } finally {
                if (btnElOpt) { btnElOpt.disabled = false; btnElOpt.textContent = btnElOpt._origText || '🔄 刷新'; }
            }
        },

        /**
         * 给定任务级 Model（select 的 value = "provider::model" 或空 AUTO），返回 {provider?, model?}。
         * 当为空时返回 {}，表示用默认执行模型（全局 llm_config.provider+model）。
         */
        parseTaskModelValue(v) {
            if (!v) return {};
            const i = v.indexOf('::');
            if (i < 0) return { model: v };
            return { provider: v.substring(0, i) || undefined, model: v.substring(i + 2) || undefined };
        },

        /**
         * 模型选择条：读取当前 Excel 侧的 LLM 配置回显到表单（内嵌面板 / 独立页通用）
         *   - provider/model/apiUrl 原样回显
         *   - apiKey 不会明文回显；只填掩码占位，用户不改动时 save 时不覆盖原值
         */
        async loadLLMBar() {
            const pv = document.getElementById('tp-llm-provider');
            const md = document.getElementById('tp-llm-model');
            const au = document.getElementById('tp-llm-apiUrl');
            const ak = document.getElementById('tp-llm-apiKey');
            const kh = document.getElementById('tp-llm-keyhint');
            // VSTO 任务窗格无 Excel 9528：从 App.config / 桥读配置，勿 fetch localhost:9528
            const isVsto = !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects);
            const applyCfg = (c) => {
                if (!c) return false;
                if (pv) pv.value = c.apiType || c.provider || '';
                if (md) md.value = c.model || '';
                if (au) au.value = c.apiUrl || '';
                if (ak) ak.value = '';
                if (kh) kh.textContent = c.apiKey ? '已填：******' : '（未填）';
                return true;
            };
            try {
                if (isVsto) {
                    if (window.App && App.config && applyCfg(App.config)) return;
                    try {
                        const bridge = (window.BridgeAccessor && BridgeAccessor.getBridge && BridgeAccessor.getBridge())
                            || (window.App && App.excelBridge);
                        if (bridge && bridge.GetAIConfig) {
                            const raw = await bridge.GetAIConfig();
                            const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
                            const cfg = (parsed && parsed.config) || parsed;
                            if (cfg && applyCfg(cfg)) return;
                        }
                    } catch (_b) { /* fallthrough */ }
                    if (kh) kh.textContent = '（使用当前聊天模型配置）';
                    return;
                }
                await getExcelToken();
                const res = await fetch(excelApiUrl(API_BASE + '/llm/config'), { method: 'GET' });
                const data = await res.json().catch(()=>({success:false}));
                if (res.ok && data && data.success) {
                    if (pv) pv.value = data.provider || '';
                    if (md) md.value = data.model || '';
                    if (au) au.value = data.apiUrl || '';
                    if (ak) ak.value = '';
                    if (kh) kh.textContent = data.hasApiKey ? ('已填：' + (data.apiKeyMasked || '******')) : '（未填）';
                    return;
                }
                if (kh) kh.textContent = '读取失败';
            } catch (e) {
                if (kh) kh.textContent = '（无法连接 Excel 服务）';
                if (!isVsto) console.warn('[TaskPlanner] loadLLMBar failed:', e.message);
            }
        },

        /**
         * 模型选择条：保存到 Excel VSTO 存储 + 立即导出明文给独立服务 +（若独立服务在线）热加载配置
         * 表单如果 apiKey 框留空 => 认为不覆盖已有密钥，后端收到空串时 SaveLLMConfig 当前会写空导致密钥丢失
         *   => 这里在前端就把空 apiKey 从 payload 里删掉，只在用户填了非空值时才覆盖
         */
        async saveLLMBar() {
            const pv = document.getElementById('tp-llm-provider');
            const md = document.getElementById('tp-llm-model');
            const au = document.getElementById('tp-llm-apiUrl');
            const ak = document.getElementById('tp-llm-apiKey');
            const sv = document.getElementById('tp-llm-save');
            const payload = {};
            if (pv && pv.value) payload.provider = pv.value.trim();
            if (md && md.value) payload.model = md.value.trim();
            if (au && au.value) payload.apiUrl = au.value.trim();
            if (ak && ak.value && ak.value.trim() !== '') payload.apiKey = ak.value.trim();

            if (!payload.provider && !payload.model && !payload.apiUrl && !payload.apiKey) {
                this._toast('请至少填一项再保存', 'error');
                return;
            }

            if (sv) { sv.disabled = true; const orig = sv.textContent; sv.textContent = '保存中…'; }
            try {
                await getExcelToken();
                const res = await fetch(excelApiUrl(API_BASE + '/llm/save-config'), {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const data = await res.json().catch(()=>({success:false}));
                if (!res.ok || !data.success) {
                    this._toast('保存失败：' + (data.error || ('HTTP ' + res.status)), 'error');
                    return;
                }

                // 成功后：如果独立服务运行中 → 热加载新配置
                let hotOk = null;
                try {
                    const r2 = await fetch((App.Endpoints ? App.Endpoints.standaloneAgent.reloadConfig : 'http://localhost:9529/reload-config'), { method: 'POST' });
                    hotOk = r2.ok;
                } catch(e) { hotOk = false; }

                const exported = (data.exported && data.exported.allPaths) ? data.exported.allPaths.length : 0;
                const msg = (exported ? ('✅ 已保存并导出到 ' + exported + ' 处') : '✅ 已保存') +
                    (hotOk === true ? '，独立服务已热加载新配置' : (hotOk === false ? '（独立服务当前不在线，下次启动自动读取）' : ''));
                this._toast(msg, 'success');

                // 清掉 apiKey 输入框（防止长时间停留在页面时误泄漏），并刷新掩码提示
                if (ak) ak.value = '';
                try { await this.loadLLMBar(); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
            } catch (e) {
                this._toast('保存失败：' + e.message, 'error');
            } finally {
                if (sv) { sv.disabled = false; sv.textContent = '保存并同步'; }
            }
        }
    };

    // 导出到全局
    window.TaskPlanner = TaskPlanner;

    // 监听任务执行事件（后端发送给前端）
    document.addEventListener('task_execution', function(e) {
        console.log('[TaskPlanner] Task execution event:', e.detail);
        // 可以在消息输入框中填入指令
        const input = document.getElementById('message-input');
        if (input && e.detail?.command) {
            input.value = e.detail.command;
            // 自动发送
            if (typeof sendMessage === 'function') {
                setTimeout(() => sendMessage(), 500);
            }
        }
    });

    // 监听来自 WebView 的消息（COM Bridge 定时器触发）
    window.addEventListener('message', function(e) {
        try {
            const data = typeof e.data === 'string' ? JSON.parse(e.data) : e.data;
            if (data.Type === 'task_execution' || data.type === 'task_execution') {
                // 设置绕过守卫标志
                window._tpBypassGuard = true;
                const event = new CustomEvent('task_execution', { detail: data });
                document.dispatchEvent(event);
                // 30 秒后清除绕过标志
                setTimeout(function() {
                    window._tpBypassGuard = false;
                }, 30000);
            }
        } catch (e) {
            // 非 JSON 消息，忽略
        }
    });

    // 前端定时检查：每 30 秒检查是否有任务需要触发（COM Bridge 定时器的降级补丁）
    setInterval(function() {
        try {
            const raw = localStorage.getItem('tp_tasks');
            if (!raw) return;
            const tasks = JSON.parse(raw);
            if (!Array.isArray(tasks) || tasks.length === 0) return;
            const now = new Date();
            const nowMin = now.getHours() * 60 + now.getMinutes();

            for (const task of tasks) {
                if (task.Enabled === false) continue;
                if (!task.Time) continue;
                const parts = task.Time.split(':');
                if (parts.length < 2) continue;
                const taskMin = parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
                if (nowMin !== taskMin) continue;

                // 检查是否今天已执行过
                const today = now.toLocaleDateString('zh-CN');
                if (task.LastRun && task.LastRun.startsWith(today)) continue;

                // 频率检查
                const dayOfWeek = now.getDay();
                if (task.Frequency === 'weekdays' && (dayOfWeek === 0 || dayOfWeek === 6)) continue;
                if (task.Frequency === 'weekly' && dayOfWeek !== 1) continue;
                if (task.Frequency === 'once' && task.LastRun) continue;

                // 时间匹配，触发执行
                console.log('[TaskPlanner] 前端定时器触发任务:', task.Name);
                const input = document.getElementById('message-input');
                const command = task.Command || '';
                if (input && command) {
                    // 设置绕过守卫标志
                    window._tpBypassGuard = true;
                    input.value = command;
                    const evt = new Event('input', { bubbles: true });
                    input.dispatchEvent(evt);
                    if (typeof sendMessage === 'function') {
                        setTimeout(function() {
                            sendMessage();
                            setTimeout(function() {
                                window._tpBypassGuard = false;
                            }, 30000);
                        }, 300);
                    }
                }

                // 更新 LastRun
                task.LastRun = new Date().toLocaleString('zh-CN');
                try {
                    localStorage.setItem('tp_tasks', JSON.stringify(tasks));
                } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TASK', '操作失败:', e && e.message); }
            }
        } catch (e) {
            console.warn('[TaskPlanner] 前端定时检查失败:', e.message);
        }
    }, 30000); // 每 30 秒检查一次

    console.log('[TaskPlanner] Plugin loaded v2026082956');
})();