/**
 * 验证 models-config.js + renderAllModelList 按平台渲染逻辑：
 *  - 本地平台 → 仅显示本地检测模型
 *  - 云端平台 → 仅显示该平台预设模型
 * 运行：node Tools/verify-model-groups.js
 */
const fs = require('fs');
const path = require('path');

const configSrc = fs.readFileSync(path.join(__dirname, '..', 'UI', 'config', 'models-config.js'), 'utf8');
const window = {};
eval(configSrc);
const PROVIDER_CONFIGS = window.MODELS_CONFIG.providerConfigs;

const appModelSrc = fs.readFileSync(path.join(__dirname, '..', 'UI', 'js', 'app-model.js'), 'utf8');
const lines = appModelSrc.split('\n');

function extractFunction(name) {
    const start = lines.findIndex(l => l.includes(`function ${name}`));
    if (start < 0) { console.error(`函数 ${name} 未找到`); process.exit(1); }
    let depth = 0, end = start;
    for (let i = start; i < lines.length; i++) {
        depth += (lines[i].match(/\{/g) || []).length - (lines[i].match(/\}/g) || []).length;
        if (depth <= 0) { end = i; break; }
    }
    return lines.slice(start, end + 1).join('\n');
}

const ChatLog = { info: () => {}, warn: () => {}, error: () => {}, debug: () => {} };
const getProviderName = (k) => ({ ollama: 'Ollama 本地', llamacpp: 'llama.cpp 本地', openai: 'OpenAI', azure: 'Azure OpenAI', wenxin: '文心一言', qianwen: '通义千问', zhipu: '智谱 AI', xunfei: '讯飞星火', moonshot: 'Moonshot', custom: '自定义' }[k] || k);
const getShortProviderName = (k) => ({ openai: 'OpenAI', azure: 'Azure', wenxin: '文心', qianwen: '通义', zhipu: '智谱', xunfei: '星火', moonshot: 'Kimi', custom: '自定义', ollama: 'Ollama', llamacpp: 'llama.cpp' }[k] || k);
const getLocalModelDisplayName = (model) => {
    const raw = (model && model.name) || '';
    if ((model && model.provider) !== 'llamacpp') return raw;
    return raw.split(/[\\/]/).pop().replace(/\.gguf$/i, '') || raw;
};

const fnSrc = [
    extractFunction('renderModelItemHtml'),
    extractFunction('renderAllModelList')
].join('\n');

const detectedModels = [
    { name: 'qwen2.5:7b', provider: 'ollama' },
    { name: 'gemma3:4b', provider: 'ollama' },
    { name: 'models\\llama-3.2-3b-q8.gguf', provider: 'llamacpp' },
    { name: 'models\\mistral-7b.gguf', provider: 'llamacpp' }
];

function run(config, label, expectCloud, expectLocal, expectNames) {
    const modelList = { innerHTML: '' };
    eval(fnSrc + `\nconfig = ${JSON.stringify(config)}; renderAllModelList(modelList, detectedModels);`);
    const html = modelList.innerHTML;
    const hasCloudGroup = /· 云端/.test(html);
    const hasLocalGroup = /本地模型/.test(html);
    const namesMatch = !expectNames || expectNames.every(n => html.includes(n));
    const notShownNames = !expectNames ? [] : expectNames.filter(n => !html.includes(n));
    const ok = (expectCloud === hasCloudGroup) && (expectLocal === hasLocalGroup) && namesMatch;
    const firstLine = html.split('\n').find(l => l.includes('云端') || l.includes('本地') || l.includes('未') || l.includes('暂无'));
    console.log(`${ok ? '✅' : '❌'} [${label}] 分组: ${firstLine ? firstLine.replace(/<[^>]+>/g, '').trim().slice(0, 45) : '(空)'}`);
    if (notShownNames.length) console.log(`    缺失模型: ${notShownNames.join(', ')}`);
    return ok;
}

console.log('========== 按平台渲染验证 ==========');
const r1 = run({ apiType: 'llamacpp', model: '' }, 'llama.cpp(本地)', false, true, ['llama-3.2-3b', 'mistral-7b']);
const r2 = run({ apiType: 'ollama', model: '' }, 'Ollama(本地)', false, true, ['qwen2.5:7b', 'gemma3:4b']);
const r3 = run({ apiType: 'openai', model: '' }, 'OpenAI(云端)', true, false, ['gpt-5.6-sol']);
const r4 = run({ apiType: 'wenxin', model: '' }, '文心一言(云端)', true, false, ['ernie-5.1']);
const r5 = run({ apiType: 'moonshot', model: '' }, 'Moonshot(云端)', true, false, ['kimi-k3']);
const r6 = run({ apiType: 'deepseek', model: '' }, 'DeepSeek(云端)', true, false, ['deepseek-v4-pro']);
const r7 = run({ apiType: 'volcengine', model: '' }, '火山引擎(云端)', true, false, ['doubao-seed-2-1-pro-260628']);
console.log('==================================');
const allOk = r1 && r2 && r3 && r4 && r5 && r6 && r7;
console.log(allOk ? '\n✅ 全部通过：本地平台按服务过滤（Ollama/llama.cpp 互不串），云端平台只显示该平台预设模型' : '\n❌ 存在失败用例');
