// ============================================================================
// 平台模型预置配置（纯数据，无逻辑）
// ----------------------------------------------------------------------------
// 各平台的 API 地址、默认模型、可用模型列表统一存放在此文件，
// 更新模型名单时只需修改本文件，无需改动任何逻辑代码。
//
// 加载方式：在 app-bootstrap.js 之前以 <script src="config/models-config.js">
// 引入，app-bootstrap.js 通过 window.MODELS_CONFIG.providerConfigs 读取。
//
// models 元素：字符串数组（仅模型名）或 { name, contextWindow } 对象数组，
// 两种格式均兼容（渲染与选择逻辑统一按 name 取值）。
// ============================================================================
window.MODELS_CONFIG = {
    providerConfigs: {
        ollama: {
            apiUrl: 'http://localhost:11434/api/chat',
            model: 'qwen2.5:7b',
            needsKey: false,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            isLocal: true,
            contextWindow: 32768,
            models: [
                { name: 'llama3.1', contextWindow: 8192 },
                { name: 'llama3.2', contextWindow: 16384 },
                { name: 'llama3.3', contextWindow: 131072 },
                { name: 'mistral', contextWindow: 32768 },
                { name: 'mixtral', contextWindow: 32768 },
                { name: 'qwen2.5:7b', contextWindow: 32768 },
                { name: 'qwen2.5:32b', contextWindow: 32768 },
                { name: 'qwen2.5:72b', contextWindow: 32768 },
                { name: 'qwen3:4b', contextWindow: 32768 },
                { name: 'qwen3:8b', contextWindow: 32768 },
                { name: 'qwen3:30b', contextWindow: 32768 },
                { name: 'deepseek-r1:1.5b', contextWindow: 16384 },
                { name: 'deepseek-r1:7b', contextWindow: 16384 },
                { name: 'deepseek-r1:14b', contextWindow: 16384 },
                { name: 'deepseek-r1:32b', contextWindow: 16384 },
                { name: 'deepseek-r1:70b', contextWindow: 16384 },
                { name: 'gemma3:4b', contextWindow: 32768 },
                { name: 'gemma3:12b', contextWindow: 32768 },
                { name: 'qwen2.5-vl', contextWindow: 32768 }
            ]
        },
        llamacpp: {
            apiUrl: 'http://localhost:8080/v1',
            model: '',
            needsKey: false,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            isLocal: true,
            contextWindow: 8192,
            models: []
        },
        openai: {
            apiUrl: 'https://api.openai.com/v1/chat/completions',
            model: 'gpt-5.6-sol',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            models: ['gpt-5.6-sol', 'gpt-5.6-terra', 'gpt-5.6-luna', 'gpt-5.5', 'gpt-5.4', 'gpt-5.4-pro', 'gpt-5.4-mini', 'gpt-5.4-nano', 'gpt-5.3-codex', 'gpt-4.1', 'gpt-4.1-mini', 'gpt-4.1-nano', 'gpt-4o', 'gpt-4o-mini', 'gpt-6-astra', 'gpt-chat-latest']
        },
        azure: {
            apiUrl: '',
            model: 'gpt-5.6-sol',
            needsKey: true,
            headerName: 'api-key',
            headerPrefix: '',
            models: ['gpt-5.6-sol', 'gpt-5.6-terra', 'gpt-5.6-luna', 'gpt-5.4', 'gpt-5.4-pro', 'gpt-5.4-mini', 'gpt-5.4-nano', 'gpt-5.3-codex', 'gpt-4.1', 'gpt-4.1-mini', 'gpt-4o', 'gpt-4o-mini', 'gpt-oss-120b', 'gpt-oss-20b'],
            placeholder: 'https://{your-resource}.openai.azure.com/openai/deployments/{deployment}/chat/completions?api-version=2024-02-01',
            description: 'Azure OpenAI 需要填写您的 Azure 服务端点 URL'
        },
        wenxin: {
            apiUrl: 'https://qianfan.baidubce.com/v2/chat/completions',
            model: 'ernie-5.1',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            models: ['ernie-5.1', 'ernie-5.0', 'ernie-x1.1', 'ernie-4.5-turbo-128k', 'ernie-4.5-turbo-32k', 'ernie-4.5-turbo-vl']
        },
        qianwen: {
            apiUrl: 'https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions',
            model: 'qwen3.8-max',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            models: ['qwen3.8-max', 'qwen3.8-flash', 'qwen3.8-max-0902', 'qwen3.7-max', 'qwen3.7-plus', 'qwen3.6-flash', 'qwen3.5-plus', 'qwen3.5-flash', 'qwen3-max', 'qwen3-coder-plus']
        },
        zhipu: {
            apiUrl: 'https://open.bigmodel.cn/api/paas/v4/chat/completions',
            model: 'glm-5.3',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            models: ['glm-5.3', 'glm-5.3-flash', 'glm-5.2', 'glm-5.1', 'glm-5.1-highspeed', 'glm-5', 'glm-5-turbo', 'glm-5v-turbo', 'glm-4.7', 'glm-4.6v']
        },
        xunfei: {
            apiUrl: 'https://spark-api-open.xf-yun.com/v1/chat/completions',
            model: 'spark-4.0-ultra',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            models: ['spark-4.0-ultra', 'spark-max', 'spark-pro', 'spark-lite', '4.0Ultra', 'pro-128k', 'max-32k', 'generalv3.5']
        },
        moonshot: {
            apiUrl: 'https://api.moonshot.cn/v1/chat/completions',
            model: 'kimi-k3',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            contextWindow: 1048576,
            models: [
                { name: 'kimi-k3', contextWindow: 1048576 },
                { name: 'kimi-k2.7-code', contextWindow: 262144 },
                { name: 'kimi-k2.7-code-highspeed', contextWindow: 262144 },
                { name: 'kimi-k2.6', contextWindow: 262144 }
            ]
        },
        deepseek: {
            apiUrl: 'https://api.deepseek.com/v1/chat/completions',
            model: 'deepseek-v4-pro',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            contextWindow: 131072,
            models: [
                { name: 'deepseek-v4-pro', contextWindow: 131072 },
                { name: 'deepseek-v4-flash', contextWindow: 131072 }
            ]
        },
        volcengine: {
            apiUrl: 'https://ark.cn-beijing.volces.com/api/v3/chat/completions',
            model: 'doubao-seed-2-1-pro-260628',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            models: [
                'doubao-seed-2-1-pro-260628',
                'doubao-seed-2-0-pro-260215',
                'doubao-seed-2-0-lite-260215',
                'doubao-seed-2-0-mini-260215',
                'doubao-seed-2-0-code-preview-260215',
                'doubao-seed-1-8-251228',
                'doubao-1-5-pro-32k-character-250228'
            ]
        },
        custom: {
            apiUrl: '',
            model: '',
            needsKey: true,
            headerName: 'Authorization',
            headerPrefix: 'Bearer ',
            models: []
        }
    }
};
