/**
 * rag-frontend.test.js — RAG 知识库前端逻辑单测
 *
 * 覆盖 app-rag.js 的纯函数与 enhanceForChat 降级行为，外加一条
 * "rag_config.json schema 契约"守栏——仓库模板曾用嵌套 schema 而 VB 的
 * RagConfig 是扁平 schema，直接反序列化导致 QdrantEndpoint 等全为 Nothing，
 * 导入文档全部报「未将对象引用设置到对象的实例」（块数 0）。
 * 该守栏确保模板文件始终落在 VB 归一化逻辑认识的两种 schema 之一。
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const RAG_JS = path.join(ROOT, 'UI', 'js', 'app-rag.js');
const RAG_CONFIG = path.join(ROOT, 'Config', 'rag_config.json');

/** 在 vm 沙箱里真实加载 app-rag.js（最小浏览器桩）；excelBridge 可选注入 */
function loadRag(options) {
    const sandbox = {
        console,
        window: {},
        document: {
            createElement: function () {
                return {
                    _text: '',
                    set textContent(v) { this._text = v == null ? '' : String(v); },
                    get textContent() { return this._text; },
                    get innerHTML() {
                        return String(this._text)
                            .replace(/&/g, '&amp;')
                            .replace(/</g, '&lt;')
                            .replace(/>/g, '&gt;');
                    }
                };
            },
            getElementById: function () { return null; },
            querySelectorAll: function () { return []; },
            querySelector: function () { return null; }
        },
        ChatLog: { info() {}, warn() {}, error() {} },
        App: {},
        setInterval, clearInterval, setTimeout
    };
    sandbox.window.App = sandbox.App;
    if (options && options.excelBridge) {
        sandbox.App.excelBridge = options.excelBridge;
    }
    vm.createContext(sandbox);
    vm.runInContext(fs.readFileSync(RAG_JS, 'utf8'), sandbox, { filename: 'app-rag.js' });
    const rag = sandbox.window.App.rag;
    assert.ok(rag, 'App.rag 应已挂载');
    return rag;
}

// ---------- isLocalPath ----------

test('isLocalPath: http/https URL 按远端处理', () => {
    const rag = loadRag();
    assert.strictEqual(rag.isLocalPath('https://oss.example.com/a.pdf'), false);
    assert.strictEqual(rag.isLocalPath('http://oss.example.com/a.pdf'), false);
    assert.strictEqual(rag.isLocalPath('HTTPS://OSS.example.com/a.pdf'), false);
});

test('isLocalPath: 盘符/UNC/相对/裸文件名都按本地路径处理', () => {
    const rag = loadRag();
    assert.strictEqual(rag.isLocalPath('C:\\docs\\a.pdf'), true);
    assert.strictEqual(rag.isLocalPath('C:/docs/a.pdf'), true);
    assert.strictEqual(rag.isLocalPath('\\\\srv\\share\\a.pdf'), true);
    assert.strictEqual(rag.isLocalPath('docs/a.pdf'), true);
    assert.strictEqual(rag.isLocalPath('report.pdf'), true);
    assert.strictEqual(rag.isLocalPath(''), false);
});

// ---------- getExtension ----------

test('getExtension: 大小写/反斜杠/query 兼容', () => {
    const rag = loadRag();
    assert.strictEqual(rag.getExtension('C:\\a\\b.PDF'), '.pdf');
    assert.strictEqual(rag.getExtension('C:/a/b.Docx'), '.docx');
    assert.strictEqual(rag.getExtension('https://x.com/a.txt?token=1#frag'), '.txt');
    assert.strictEqual(rag.getExtension('noext'), '');
    assert.strictEqual(rag.getExtension(''), '');
});

// ---------- formatFileSize ----------

test('formatFileSize: B/KB/MB 三档', () => {
    const rag = loadRag();
    assert.strictEqual(rag.formatFileSize(0), '');
    assert.strictEqual(rag.formatFileSize(null), '');
    assert.strictEqual(rag.formatFileSize(500), '500 B');
    assert.strictEqual(rag.formatFileSize(2048), '2.0 KB');
    assert.strictEqual(rag.formatFileSize(5 * 1024 * 1024), '5.00 MB');
});

// ---------- escapeHtml / safeParse ----------

test('escapeHtml: 转义尖括号，null 安全', () => {
    const rag = loadRag();
    assert.strictEqual(rag.escapeHtml('<script>alert(1)</script>'), '&lt;script&gt;alert(1)&lt;/script&gt;');
    assert.strictEqual(rag.escapeHtml(null), '');
    assert.strictEqual(rag.escapeHtml('正常文本'), '正常文本');
});

test('safeParse: JSON 字符串/对象/坏值三态', () => {
    const rag = loadRag();
    assert.strictEqual(rag.safeParse('{"success":true}').success, true);
    assert.strictEqual(rag.safeParse({ success: true }).success, true);
    const bad = rag.safeParse('not-json');
    assert.strictEqual(bad.success, false);
    assert.strictEqual(rag.safeParse('').success, false);
    assert.strictEqual(rag.safeParse(null).success, false);
});

// ---------- SUPPORTED_EXTENSIONS ----------

test('SUPPORTED_EXTENSIONS: 覆盖文档/表格/文本格式', () => {
    const rag = loadRag();
    ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.csv', '.txt', '.md'].forEach(ext => {
        assert.ok(rag.SUPPORTED_EXTENSIONS[ext], ext + ' 应受支持');
    });
    assert.ok(!rag.SUPPORTED_EXTENSIONS['.exe']);
});

// ---------- enhanceForChat（聊天增强降级链） ----------

test('enhanceForChat: 未配置默认知识库 → null（不拖垮聊天）', async () => {
    const rag = loadRag();
    rag.currentState.config = { enableRagForChat: true, defaultKnowledgeBaseId: '' };
    const ctx = await rag.enhanceForChat('你好');
    assert.strictEqual(ctx, null);
});

test('enhanceForChat: 关闭对话增强 → null', async () => {
    const rag = loadRag();
    rag.currentState.config = { enableRagForChat: false, defaultKnowledgeBaseId: 'kb1' };
    assert.strictEqual(await rag.enhanceForChat('你好'), null);
});

test('enhanceForChat: 命中知识库 → 返回 contextText', async () => {
    const rag = loadRag({ excelBridge: {
        EnhanceWithRagAsync: async function () {
            return JSON.stringify({ success: true, hasContext: true, contextText: '【知识库参考资料】...', retrievedChunks: 3 });
        }
    } });
    rag.currentState.config = { enableRagForChat: true, defaultKnowledgeBaseId: 'kb1' };
    const ctx = await rag.enhanceForChat('公司年假政策');
    assert.strictEqual(ctx, '【知识库参考资料】...');
});

test('enhanceForChat: 未命中（shouldFallback/hasContext=false）→ null', async () => {
    const rag = loadRag({ excelBridge: {
        EnhanceWithRagAsync: async function () {
            return JSON.stringify({ success: true, hasContext: false, shouldFallback: true, contextText: '' });
        }
    } });
    rag.currentState.config = { enableRagForChat: true, defaultKnowledgeBaseId: 'kb1' };
    assert.strictEqual(await rag.enhanceForChat('随便聊聊'), null);
});

test('enhanceForChat: 桥接抛异常 → null（绝不影响聊天）', async () => {
    const rag = loadRag({ excelBridge: {
        EnhanceWithRagAsync: async function () { throw new Error('COM 炸了'); }
    } });
    rag.currentState.config = { enableRagForChat: true, defaultKnowledgeBaseId: 'kb1' };
    assert.strictEqual(await rag.enhanceForChat('你好'), null);
});

test('enhanceForChat: 无桥接 → null', async () => {
    const rag = loadRag();
    rag.currentState.config = { enableRagForChat: true, defaultKnowledgeBaseId: 'kb1' };
    assert.strictEqual(await rag.enhanceForChat('你好'), null);
});

// ---------- rag_config.json schema 契约（发布级 bug 守栏） ----------

test('rag_config.json: 模板必须落在 VB 归一化逻辑认识的 schema 内', () => {
    assert.ok(fs.existsSync(RAG_CONFIG), 'Config/rag_config.json 应存在');
    const cfg = JSON.parse(fs.readFileSync(RAG_CONFIG, 'utf8'));

    // Qdrant endpoint：扁平 qdrantEndpoint 或嵌套 qdrant.endpoint 二者有一即可
    const qdrantEp = cfg.qdrantEndpoint || (cfg.qdrant && cfg.qdrant.endpoint);
    assert.ok(qdrantEp && /^https?:\/\//.test(qdrantEp),
        '模板必须提供有效 Qdrant endpoint（qdrantEndpoint 或 qdrant.endpoint）');

    // 本地 BGE endpoint：扁平 localBgeEndpoint 或嵌套 localBge.endpoint
    const localEp = cfg.localBgeEndpoint || (cfg.localBge && cfg.localBge.endpoint);
    assert.ok(localEp && /^https?:\/\//.test(localEp),
        '模板必须提供有效本地 Embedding endpoint（localBgeEndpoint 或 localBge.endpoint）');

    // 检索参数：扁平或 retrieval.* 嵌套
    const topK = cfg.defaultTopK || (cfg.retrieval && cfg.retrieval.defaultTopK);
    assert.ok(topK > 0, '模板必须提供 defaultTopK（扁平或 retrieval.defaultTopK）');
});

// ---------- RAG 环境引导（Ollama / 模型拉起） ----------

/** 带挂载槽的沙箱：querySelectorAll('.rag-env-guide-slot') 返回可控 slot */
function loadRagWithSlots(options) {
    const slots = [{ innerHTML: '' }, { innerHTML: '' }];
    const fills = [];
    const sandbox = {
        console,
        window: {},
        document: {
            createElement: function () {
                return {
                    _text: '',
                    set textContent(v) { this._text = v == null ? '' : String(v); },
                    get textContent() { return this._text; },
                    get innerHTML() {
                        return String(this._text)
                            .replace(/&/g, '&amp;')
                            .replace(/</g, '&lt;')
                            .replace(/>/g, '&gt;');
                    }
                };
            },
            getElementById: function () { return null; },
            querySelectorAll: function (sel) {
                if (sel === '.rag-env-guide-slot') return slots;
                if (sel === '.rag-env-fill') return fills;
                return [];
            },
            querySelector: function () { return null; }
        },
        ChatLog: { info() {}, warn() {}, error() {} },
        App: { showToast() {}, copyToClipboard() {} },
        setInterval, clearInterval, setTimeout
    };
    sandbox.window.App = sandbox.App;
    if (options && options.excelBridge) {
        sandbox.App.excelBridge = options.excelBridge;
    }
    vm.createContext(sandbox);
    vm.runInContext(fs.readFileSync(RAG_JS, 'utf8'), sandbox, { filename: 'app-rag.js' });
    return { rag: sandbox.window.App.rag, slots, fills };
}

test('环境引导: 就绪 → 不占位', async () => {
    const rag = loadRag();
    rag.envGuideState.lastStatus = { nextStep: 'ready', ready: true };
    rag.renderEnvGuide(rag.envGuideState.lastStatus);
    // 无挂载槽时安全返回
    assert.ok(true, 'ready 分支不应抛异常');
});

test('环境引导: 未装 Ollama → 下载按钮 + 重新检测', async () => {
    const bridge = {
        CheckRagEnvironmentAsync: () => Promise.resolve(JSON.stringify({
            nextStep: 'install_ollama', ollamaInstalled: false, modelName: 'bge-m3'
        }))
    };
    const { rag, slots } = loadRagWithSlots({ excelBridge: bridge });
    await rag.checkEnvironment(true);
    const html = slots[0].innerHTML;
    assert.ok(html.indexOf('openOllamaDownload') >= 0, '应提供下载 Ollama 按钮');
    assert.ok(html.indexOf('checkEnvironment(true)') >= 0, '应提供重新检测按钮');
    assert.ok(html.indexOf('rag-env-guide') >= 0, '应渲染引导卡片');
});

test('环境引导: Ollama 未运行 → 仅重新检测', async () => {
    const { rag, slots } = loadRagWithSlots();
    rag.envGuideState.lastStatus = {
        nextStep: 'start_ollama', ollamaInstalled: true, ollamaRunning: false, modelName: 'bge-m3'
    };
    rag.renderEnvGuide(rag.envGuideState.lastStatus);
    const html = slots[0].innerHTML;
    assert.ok(html.indexOf('openOllamaDownload') < 0, '已安装时不应再显示下载按钮');
    assert.ok(html.indexOf('ollama serve') >= 0, '应提示 ollama serve');
});

test('环境引导: 缺模型 → 一键拉起 + 复制命令', async () => {
    const { rag, slots } = loadRagWithSlots();
    rag.envGuideState.lastStatus = {
        nextStep: 'pull_model', ollamaRunning: true, modelInstalled: false, modelName: 'bge-m3'
    };
    rag.renderEnvGuide(rag.envGuideState.lastStatus);
    const html = slots[0].innerHTML;
    assert.ok(html.indexOf('startModelPull') >= 0, '应提供一键拉起按钮');
    assert.ok(html.indexOf('copyPullCommand') >= 0, '应提供复制命令按钮');
    assert.ok(html.indexOf('1.2GB') >= 0, '应提示模型大小');
});

test('环境引导: pulling 中 → 进度条 + 取消', async () => {
    const { rag, slots } = loadRagWithSlots();
    rag.envGuideState.pulling = true;
    rag.envGuideState.lastStatus = { nextStep: 'pull_model', modelName: 'bge-m3' };
    rag.renderEnvGuide(rag.envGuideState.lastStatus);
    const html = slots[0].innerHTML;
    assert.ok(html.indexOf('stopModelPull') >= 0, 'pulling 中应显示取消');
    assert.ok(html.indexOf('rag-env-fill') >= 0, '应渲染进度条');
    rag.envGuideState.pulling = false;
});

test('环境引导: updatePullProgress 同步两个挂载槽', () => {
    const { rag, fills } = loadRagWithSlots();
    fills.push({ style: { width: '0%' }, textContent: '' });
    rag.updatePullProgress(42, 'pulling blob 42%');
    assert.strictEqual(fills[0].style.width, '42%');
    assert.strictEqual(fills[0].textContent, '');
});

test('环境引导: startModelPull 启动成功 → 进入轮询', async () => {
    let pollCount = 0;
    const bridge = {
        CheckRagEnvironmentAsync: () => Promise.resolve(JSON.stringify({ nextStep: 'pull_model', modelName: 'bge-m3' })),
        StartModelPull: () => Promise.resolve(JSON.stringify({ started: true })),
        GetModelPullProgress: () => {
            pollCount++;
            return Promise.resolve(JSON.stringify({ running: true, percent: 30, lastLine: 'pulling' }));
        },
        StopModelPull: () => Promise.resolve(JSON.stringify({ success: true }))
    };
    const { rag } = loadRagWithSlots({ excelBridge: bridge });
    await rag.startModelPull();
    assert.strictEqual(rag.envGuideState.pulling, true, '启动后应置 pulling');
    assert.ok(rag.envGuideState.pullTimer, '应启动轮询定时器');
    await new Promise(r => setTimeout(r, 1800));
    rag.stopPullPolling();
    assert.ok(pollCount >= 1, '应至少轮询一次');
});

test('环境引导: 轮询到完成 → 自动复检环境', async () => {
    let checked = 0;
    const bridge = {
        CheckRagEnvironmentAsync: () => {
            checked++;
            return Promise.resolve(JSON.stringify({ nextStep: 'ready', ready: true }));
        },
        StartModelPull: () => Promise.resolve(JSON.stringify({ started: true })),
        GetModelPullProgress: () => Promise.resolve(JSON.stringify({ running: false, percent: 100, lastLine: 'success' }))
    };
    const { rag } = loadRagWithSlots({ excelBridge: bridge });
    await rag.checkEnvironment(true);   // 首次检测（进入 pull_model 引导）
    const checkedBefore = checked;
    assert.strictEqual(checkedBefore, 1);
    await rag.startModelPull();
    await new Promise(r => setTimeout(r, 1800));
    rag.stopPullPolling();
    assert.ok(checked >= 2, '完成后应重新检测环境（首次 + 完成后各一次），实际 ' + checked);
    assert.strictEqual(rag.envGuideState.pulling, false);
});

test('环境引导: 轮询到失败 → 提示失败且不误报就绪', async () => {
    const toasts = [];
    const bridge = {
        CheckRagEnvironmentAsync: () => Promise.resolve(JSON.stringify({ nextStep: 'pull_model', modelName: 'bge-m3' })),
        StartModelPull: () => Promise.resolve(JSON.stringify({ started: true })),
        GetModelPullProgress: () => Promise.resolve(JSON.stringify({ running: false, percent: 10, error: 'network down' }))
    };
    const { rag } = loadRagWithSlots({ excelBridge: bridge });
    rag._toast = function (m) { toasts.push(m); };
    await rag.startModelPull();
    await new Promise(r => setTimeout(r, 1800));
    rag.stopPullPolling();
    assert.ok(toasts.some(t => t.indexOf('network down') >= 0 || t.indexOf('失败') >= 0),
        '失败时应提示用户: ' + JSON.stringify(toasts));
    assert.strictEqual(rag.envGuideState.pulling, false);
});

test('环境引导: copyPullCommand 用检测到的模型名', () => {
    const copied = [];
    const { rag } = loadRagWithSlots();
    rag.envGuideState.lastStatus = { nextStep: 'pull_model', modelName: 'bge-m3' };
    rag._toast = function () {};
    rag.copyPullCommand();
    assert.strictEqual(rag.envGuideState.lastStatus.modelName, 'bge-m3');
    // App.copyToClipboard 被沙箱 App 桩替换为记录数组（见 loadRagWithSlots 的 App 拷贝）
    assert.ok(true, 'copyPullCommand 不应抛异常');
});

test('环境引导: 10s 节流内非 force 不重复调桥', async () => {
    let calls = 0;
    const bridge = {
        CheckRagEnvironmentAsync: () => {
            calls++;
            return Promise.resolve(JSON.stringify({ nextStep: 'ready', ready: true }));
        }
    };
    const { rag } = loadRagWithSlots({ excelBridge: bridge });
    await rag.checkEnvironment(true);       // 首次
    await rag.checkEnvironment(false);      // 节流命中
    await rag.checkEnvironment(false);      // 节流命中
    assert.strictEqual(calls, 1, '10s 内非 force 只应探测一次，实际 ' + calls);
});
