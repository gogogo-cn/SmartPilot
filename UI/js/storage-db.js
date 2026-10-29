/**
 * ============================================================
 * storage-db.js — Storage Database
 * ============================================================
 * 功能：存储数据库模块，封装 localStorage / IndexedDB 操作，提供键值与对象存储能力。
 *
 * 核心类/函数：
 *   - getItem() — 读取数据
 *   - setItem() — 写入数据
 *   - removeItem() — 删除数据
 *   - clear() — 清空存储
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    const DB_NAME = 'ExcelAIVSTOAddIn_DB';
    const DB_VERSION = 1;
    const STORE_NAME = 'kv_store';

    let _db = null;
    let _initPromise = null;
    let _available = null; // null=未检测, true=可用, false=不可用

    // ==================== 可用性检测 ====================

    function checkAvailable() {
        if (_available !== null) return _available;
        try {
            if (!window.indexedDB) {
                ChatLog.info('STORAGE', '[StorageDB] indexedDB 不支持');
                _available = false;
                return false;
            }
            _available = true;
            return true;
        } catch (e) {
            ChatLog.warn('STORAGE', '[StorageDB] 检测 indexedDB 失败:', e.message);
            _available = false;
            return false;
        }
    }

    // ==================== 数据库初始化 ====================

    /**
     * 打开/升级数据库（幂等，多次调用安全）
     */
    function openDB() {
        if (_db) return Promise.resolve(_db);
        if (_initPromise) return _initPromise;

        _initPromise = new Promise(function(resolve, reject) {
            const request = window.indexedDB.open(DB_NAME, DB_VERSION);

            request.onupgradeneeded = function(event) {
                ChatLog.info('STORAGE', '[StorageDB] 数据库升级中, version:', event.oldVersion, '→', DB_VERSION);
                const db = event.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME, { keyPath: 'key' });
                    ChatLog.info('STORAGE', '[StorageDB] 创建对象仓库:', STORE_NAME);
                }
            };

            request.onsuccess = function(event) {
                _db = event.target.result;

                // 监听连接关闭，清理引用
                _db.onclose = function() {
                    ChatLog.warn('STORAGE', '[StorageDB] 数据库连接已关闭');
                    _db = null;
                    _initPromise = null;
                };

                _db.onerror = function(event) {
                    ChatLog.error('STORAGE', '[StorageDB] 数据库错误:', event.target.error);
                };

                ChatLog.info('STORAGE', '[StorageDB] 数据库打开成功');
                resolve(_db);
            };

            request.onerror = function(event) {
                ChatLog.error('STORAGE', '[StorageDB] 数据库打开失败:', event.target.error);
                _available = false;
                _initPromise = null;
                reject(event.target.error);
            };

            request.onblocked = function() {
                ChatLog.warn('STORAGE', '[StorageDB] 数据库打开被阻塞（可能有其他标签页打开）');
            };
        });

        return _initPromise;
    }

    // ==================== CRUD 操作 ====================

    /**
     * 获取值
     * @param {string} key
     * @returns {Promise<any>} 解析后的值，不存在则返回 undefined
     */
    function get(key) {
        if (!checkAvailable()) return Promise.resolve(undefined);

        return openDB().then(function(db) {
            return new Promise(function(resolve, reject) {
                const tx = db.transaction([STORE_NAME], 'readonly');
                const store = tx.objectStore(STORE_NAME);
                const request = store.get(key);

                request.onsuccess = function() {
                    const result = request.result;
                    resolve(result ? result.value : undefined);
                };

                request.onerror = function() {
                    ChatLog.warn('STORAGE', '[StorageDB] 读取失败:', key, request.error);
                    reject(request.error);
                };
            });
        }).catch(function(err) {
            ChatLog.warn('STORAGE', '[StorageDB] get 异常:', key, err);
            return undefined;
        });
    }

    /**
     * 设置值
     * @param {string} key
     * @param {any} value
     * @returns {Promise<boolean>}
     */
    function set(key, value) {
        if (!checkAvailable()) return Promise.resolve(false);

        return openDB().then(function(db) {
            return new Promise(function(resolve, reject) {
                const tx = db.transaction([STORE_NAME], 'readwrite');
                const store = tx.objectStore(STORE_NAME);

                // 使用 put 实现覆盖写入
                const request = store.put({ key: key, value: value });

                request.onsuccess = function() {
                    resolve(true);
                };

                request.onerror = function() {
                    ChatLog.warn('STORAGE', '[StorageDB] 写入失败:', key, request.error);
                    reject(request.error);
                };
            });
        }).catch(function(err) {
            ChatLog.warn('STORAGE', '[StorageDB] set 异常:', key, err);
            return false;
        });
    }

    /**
     * 删除键
     * @param {string} key
     * @returns {Promise<boolean>}
     */
    function remove(key) {
        if (!checkAvailable()) return Promise.resolve(false);

        return openDB().then(function(db) {
            return new Promise(function(resolve, reject) {
                const tx = db.transaction([STORE_NAME], 'readwrite');
                const store = tx.objectStore(STORE_NAME);
                const request = store.delete(key);

                request.onsuccess = function() {
                    resolve(true);
                };

                request.onerror = function() {
                    ChatLog.warn('STORAGE', '[StorageDB] 删除失败:', key, request.error);
                    reject(request.error);
                };
            });
        }).catch(function(err) {
            ChatLog.warn('STORAGE', '[StorageDB] remove 异常:', key, err);
            return false;
        });
    }

    /**
     * 清空所有数据
     * @returns {Promise<boolean>}
     */
    function clear() {
        if (!checkAvailable()) return Promise.resolve(false);

        return openDB().then(function(db) {
            return new Promise(function(resolve, reject) {
                const tx = db.transaction([STORE_NAME], 'readwrite');
                const store = tx.objectStore(STORE_NAME);
                const request = store.clear();

                request.onsuccess = function() {
                    ChatLog.info('STORAGE', '[StorageDB] 已清空所有数据');
                    resolve(true);
                };

                request.onerror = function() {
                    ChatLog.warn('STORAGE', '[StorageDB] 清空失败:', request.error);
                    reject(request.error);
                };
            });
        }).catch(function(err) {
            ChatLog.warn('STORAGE', '[StorageDB] clear 异常:', err);
            return false;
        });
    }

    /**
     * 获取存储统计信息
     * @returns {Promise<{usedBytes: number, itemCount: number}>}
     */
    function getStats() {
        if (!checkAvailable() || !_db) {
            return Promise.resolve({ usedBytes: 0, itemCount: 0 });
        }

        // 使用浏览器估算API（如果可用）
        if (navigator.storage && navigator.storage.estimate) {
            return navigator.storage.estimate().then(function(estimate) {
                return {
                    usedBytes: estimate.usage || 0,
                    quota: estimate.quota || 0,
                    usagePercent: estimate.quota ? ((estimate.usage / estimate.quota) * 100).toFixed(1) : 0
                };
            }).catch(function() {
                return { usedBytes: 0, quota: 0, usagePercent: 0 };
            });
        }
        return Promise.resolve({ usedBytes: 0, quota: 0, usagePercent: 0 });
    }

    // ==================== 导出 ====================

    window.StorageDB = {
        DB_NAME: DB_NAME,
        STORE_NAME: STORE_NAME,

        /** 检测 IndexedDB 是否可用 */
        isAvailable: checkAvailable,

        /** 预初始化数据库（建议在应用启动时调用） */
        init: function() {
            if (!checkAvailable()) return Promise.resolve(false);
            return openDB().then(function() {
                ChatLog.info('STORAGE', '[StorageDB] 初始化完成');
                return true;
            }).catch(function(err) {
                ChatLog.warn('STORAGE', '[StorageDB] 初始化失败:', err);
                return false;
            });
        },

        get: get,
        set: set,
        remove: remove,
        clear: clear,
        getStats: getStats
    };

    ChatLog.info('STORAGE', '[StorageDB] 模块加载完成, available:', checkAvailable());

})();
