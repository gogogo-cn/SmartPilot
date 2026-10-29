/**
 * ============================================================
 * pdf-ocr-runner.js — PDF 按页 OCR 任务调度器（实时 + 断点续传）
 * ============================================================
 * 功能：
 *   1. 将 PDF 全部识别拆分为逐页任务，每页识别完立即回调。
 *   2. 支持暂停、恢复、取消；状态通过 localStorage 持久化。
 *   3. 页面/Excel 关闭或崩溃后，可基于文件指纹恢复未完成任务。
 *   4. 每完成一页即保存进度，避免单页失败导致整批丢失。
 *
 * 依赖：
 *   - 外部提供 renderPage(pageNum) -> Promise<base64>
 *   - 外部提供 recognizePage(base64, fileName) -> Promise<text>
 * ============================================================
 */
(function (global) {
    'use strict';

    const VERSION = 'v2026072422';
    const STORAGE_PREFIX = 'pdf_ocr_job:';
    const STORAGE_INDEX_KEY = 'pdf_ocr_job_index';

    function generateId() {
        return Date.now().toString(36) + Math.random().toString(36).substr(2, 6);
    }

    function stableFileKey(fileName, fileSize, totalPages) {
        const safeName = String(fileName || '').replace(/[^\w\u4e00-\u9fa5.-]/g, '_');
        return `${safeName}:${fileSize}:${totalPages}`;
    }

    function storageKey(jobId) {
        return STORAGE_PREFIX + jobId;
    }

    function readStorage(key) {
        try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : null;
        } catch (e) {
            console.warn('[PdfOcrRunner] 读取 storage 失败:', e);
            return null;
        }
    }

    function writeStorage(key, value) {
        try {
            localStorage.setItem(key, JSON.stringify(value));
            return true;
        } catch (e) {
            console.warn('[PdfOcrRunner] 写入 storage 失败:', e);
            return false;
        }
    }

    function removeStorage(key) {
        try {
            localStorage.removeItem(key);
            return true;
        } catch (e) {
            return false;
        }
    }

    function readIndex() {
        const data = readStorage(STORAGE_INDEX_KEY);
        return Array.isArray(data) ? data : [];
    }

    function writeIndex(index) {
        writeStorage(STORAGE_INDEX_KEY, index.slice(0, 20));
    }

    function addToIndex(jobId, fileKey) {
        const index = readIndex();
        const filtered = index.filter(item => item.jobId !== jobId);
        filtered.unshift({ jobId, fileKey, createdAt: Date.now() });
        writeIndex(filtered);
    }

    function removeFromIndex(jobId) {
        const index = readIndex().filter(item => item.jobId !== jobId);
        writeIndex(index);
    }

    function wait(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    class PdfOcrRunner {
        constructor(options = {}) {
            this.options = {
                onStatusChange: options.onStatusChange || function () {},
                onProgress: options.onProgress || function () {},
                onPageComplete: options.onPageComplete || function () {},
                onPageError: options.onPageError || function () {},
                onComplete: options.onComplete || function () {},
                renderPage: options.renderPage || null,
                recognizePage: options.recognizePage || null
            };
            this._runningJobs = new Set();
        }

        /**
         * 创建一个新的 PDF OCR 任务（不启动）
         * @param {File} file - PDF 文件对象
         * @param {number} totalPages - 总页数
         * @param {object} settings - 附加设置（engine、language、format、ocrAllPages 等）
         * @returns {string} jobId
         */
        createJob(file, totalPages, settings = {}) {
            const jobId = generateId();
            const fileKey = stableFileKey(file.name, file.size, totalPages);
            const state = {
                jobId,
                fileKey,
                fileName: file.name,
                fileSize: file.size,
                totalPages,
                currentPage: 0,
                currentPageProcessing: 0,
                results: [],
                failedPages: [],
                status: 'idle',
                settings,
                createdAt: Date.now(),
                updatedAt: Date.now(),
                version: VERSION
            };
            writeStorage(storageKey(jobId), state);
            addToIndex(jobId, fileKey);
            return jobId;
        }

        /**
         * 获取任务状态（始终从 storage 读取，保证恢复一致性）
         */
        getJobState(jobId) {
            return readStorage(storageKey(jobId));
        }

        /**
         * 更新任务状态字段并持久化
         */
        _updateState(jobId, patch) {
            const state = this.getJobState(jobId);
            if (!state) return null;
            Object.assign(state, patch, { updatedAt: Date.now() });
            writeStorage(storageKey(jobId), state);
            return state;
        }

        /**
         * 列出所有未完成的任务（按创建时间倒序）
         */
        listIncompleteJobs() {
            const index = readIndex();
            const jobs = [];
            for (const item of index) {
                const state = this.getJobState(item.jobId);
                if (!state) {
                    removeFromIndex(item.jobId);
                    continue;
                }
                if (state.status === 'completed' || state.status === 'cancelled') continue;
                jobs.push(state);
            }
            return jobs.sort((a, b) => b.createdAt - a.createdAt);
        }

        /**
         * 根据文件指纹查找可恢复的任务
         */
        findResumableJob(fileName, fileSize, totalPages) {
            const fileKey = stableFileKey(fileName, fileSize, totalPages);
            const incomplete = this.listIncompleteJobs();
            return incomplete.find(job => job.fileKey === fileKey) || null;
        }

        /**
         * 启动或继续任务
         * @param {string} jobId
         */
        async startJob(jobId) {
            let state = this.getJobState(jobId);
            if (!state) throw new Error('任务不存在: ' + jobId);
            if (state.status === 'completed' || state.status === 'cancelled') {
                throw new Error('任务已结束，无法继续: ' + state.status);
            }
            if (this._runningJobs.has(jobId)) {
                console.log('[PdfOcrRunner] 任务已在运行:', jobId);
                return state;
            }

            this._runningJobs.add(jobId);
            state.status = 'running';
            state.startedAt = state.startedAt || Date.now();
            this._persistAndNotify(state);

            try {
                await this._runLoop(jobId);
            } catch (e) {
                console.error('[PdfOcrRunner] 任务异常:', e);
                state = this._updateState(jobId, { status: 'failed', error: e.message });
                this._notifyStatusChange(state);
            } finally {
                this._runningJobs.delete(jobId);
            }
            return this.getJobState(jobId);
        }

        /**
         * 暂停任务（异步检查，runLoop 会在下轮读取到 paused 状态）
         */
        pauseJob(jobId) {
            const state = this.getJobState(jobId);
            if (!state || state.status !== 'running') return false;
            this._updateState(jobId, { status: 'paused' });
            this._notifyStatusChange(this.getJobState(jobId));
            return true;
        }

        /**
         * 恢复任务
         */
        resumeJob(jobId) {
            const state = this.getJobState(jobId);
            if (!state || (state.status !== 'paused' && state.status !== 'idle' && state.status !== 'running')) {
                return false;
            }
            if (this._runningJobs.has(jobId)) return true;
            this.startJob(jobId);
            return true;
        }

        /**
         * 取消任务
         */
        cancelJob(jobId) {
            const state = this.getJobState(jobId);
            if (!state) return false;
            this._updateState(jobId, { status: 'cancelled' });
            this._runningJobs.delete(jobId);
            this._notifyStatusChange(this.getJobState(jobId));
            return true;
        }

        /**
         * 完成任务后清理 storage（保留结果摘要一段时间）
         */
        disposeJob(jobId) {
            removeStorage(storageKey(jobId));
            removeFromIndex(jobId);
            this._runningJobs.delete(jobId);
        }

        /**
         * 主循环：逐页渲染 + 识别
         */
        async _runLoop(jobId) {
            let state = this.getJobState(jobId);
            if (!state) return;

            const total = state.totalPages;
            const startFrom = Math.max(0, Math.min(state.currentPage, total - 1));

            for (let pageNum = startFrom + 1; pageNum <= total; pageNum++) {
                // 每次迭代重新读取状态，以响应暂停/取消
                state = this.getJobState(jobId);
                if (!state) return;

                // 处理暂停
                while (state && state.status === 'paused') {
                    this._notifyStatusChange(state);
                    await wait(500);
                    state = this.getJobState(jobId);
                }

                if (!state || state.status === 'cancelled') {
                    this._notifyStatusChange(state);
                    return;
                }

                state.currentPageProcessing = pageNum;
                state.status = 'running';
                this._persistAndNotify(state);

                let pageText = '';
                let pageError = null;

                try {
                    if (typeof this.options.renderPage !== 'function') {
                        throw new Error('未提供 renderPage 回调');
                    }
                    const base64 = await this.options.renderPage(pageNum);
                    if (!base64) throw new Error('第 ' + pageNum + ' 页渲染为空');

                    if (typeof this.options.recognizePage !== 'function') {
                        throw new Error('未提供 recognizePage 回调');
                    }
                    pageText = await this.options.recognizePage(base64, `${state.fileName}_page${pageNum}.png`);
                } catch (err) {
                    pageError = err;
                    console.error(`[PdfOcrRunner] 第 ${pageNum}/${total} 页识别失败:`, err);
                }

                // 重新读取状态（避免回调中修改被覆盖）
                state = this.getJobState(jobId);
                if (!state || state.status === 'cancelled') return;

                if (pageError) {
                    state.failedPages.push({ page: pageNum, error: pageError.message || '未知错误' });
                    this.options.onPageError(pageNum, pageError, state);
                } else {
                    state.results.push({ page: pageNum, text: pageText });
                    state.currentPage = pageNum;
                    this.options.onPageComplete(pageNum, pageText, state);
                }

                this._persistAndNotify(state);

                // 页间短暂休息，避免连续压垮后端
                if (pageNum < total) {
                    await wait(state.settings?.pageDelayMs || 200);
                }
            }

            state = this.getJobState(jobId);
            if (!state) return;
            state.status = state.failedPages.length > 0 ? 'completed-with-errors' : 'completed';
            state.currentPageProcessing = 0;
            state.completedAt = Date.now();
            this._persistAndNotify(state);
            this.options.onComplete(state);
        }

        _persistAndNotify(state) {
            writeStorage(storageKey(state.jobId), state);
            this._notifyStatusChange(state);
            this.options.onProgress(state);
        }

        _notifyStatusChange(state) {
            try {
                this.options.onStatusChange(state);
            } catch (e) {
                console.warn('[PdfOcrRunner] onStatusChange 回调异常:', e);
            }
        }
    }

    global.PdfOcrRunner = PdfOcrRunner;
    console.log('[PdfOcrRunner] 已加载', VERSION);
})(window);
