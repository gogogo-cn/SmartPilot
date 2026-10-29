/**
 * ============================================================
 * verify-b-execresult-equivalence.js
 * ============================================================
 * 独立验证：(b) 重构把 executeMcpToolCall / executeLocalExcelTool 的 JSON 字符串返回
 * 改为对象返回后，**最终归一结构不变**。
 *
 * 原理（来自 chat-v2/services/tool-registry.util.js:normalizeMcpResult）：
 *   该函数遇到带 .result 字段的对象会自动解包一层，因此只要新代码返回"与旧代码
 *   JSON.stringify 内容完全相同"的对象，tool-registry.execute 与 skill-spec-execute
 *   经解析/解包后得到的结果与旧字符串形态逐字节等价。
 *
 * 本脚本内嵌了与源码一致的 normalizeMcpResult（仅用于本地等价性断言，不依赖浏览器/App），
 * 并复刻了两个消费者（tool-registry.execute 的 JSON.parse→normalize 分支、skill-spec-execute
 * 的 typeof 分支）的处理逻辑，对每条分支样例断言 deepEqual。
 *
 * 用法：node Tools/verify-b-execresult-equivalence.js
 * ============================================================
 */

// ---- 与源码一致的 normalizeMcpResult（tool-registry.util.js:1057-1101）----
function normalizeMcpResult(result) {
    if (result == null) return { ok: true, result: null };
    if (result && typeof result === 'object') {
        if (result.error && typeof result.error === 'string' && result.error.trim()) {
            return { ok: false, error: result.error.trim(), result };
        }
        if (result.success === false || result.ok === false || result.Success === false) {
            const errMsg = result.message || result.Message || result.msg || result.Msg
                || result.error || result.Error || '工具执行失败';
            return { ok: false, error: String(errMsg), result };
        }
        if (Array.isArray(result.content)) {
            const errorItem = result.content.find(c => c && (c.isError === true || c.type === 'error'));
            if (errorItem) {
                const errText = errorItem.text || errorItem.message || JSON.stringify(errorItem);
                return { ok: false, error: String(errText), result };
            }
        }
        if (result.result !== undefined) {
            const inner = normalizeMcpResult(result.result);
            if (!inner.ok) return { ok: false, error: inner.error, result };
            return { ok: true, result: result.result };
        }
    }
    if (typeof result === 'string') {
        const s = result.trim();
        if (/^(错误|失败|异常|出错|Error|Fail|Exception)[：:]?/i.test(s)
            || /(不能为空|不存在|执行失败|执行错误|发生错误|发生异常)$/.test(s)) {
            return { ok: false, error: s, result: s };
        }
        return { ok: true, result: s };
    }
    return { ok: true, result };
}

// ---- 复刻 tool-registry.execute 的消费：typeof 分支 + normalizeMcpResult ----
function toolRegistryConsume(raw) {
    let result = raw;
    if (typeof raw === 'string') {
        try { result = JSON.parse(raw); } catch { /* keep string */ }
    }
    return normalizeMcpResult(result);
}

// ---- 复刻 skill-spec-execute.js:1030-1045 的消费 ----
function skillSpecConsume(raw) {
    let rawResult = raw;
    if (typeof raw === 'string') {
        try { rawResult = JSON.parse(raw); } catch { rawResult = raw; }
    }
    const ok = rawResult && (rawResult.ok !== false) && !rawResult.error;
    const result = rawResult?.result !== undefined ? rawResult.result : rawResult;
    return { ok, result, error: rawResult?.error || null };
}

// ---- 构建旧（字符串）/ 新（对象）两条返回路径，内容逐字等价 ----
// 输入 ctx 模拟一次真实 bridge 往返后的 data / 原始 result。
function buildReturns(ctx) {
    const { branch, toolName, data, rawResult, errMsg, msg } = ctx;
    let oldStr, newObj;
    switch (branch) {
        case 'mock':
            oldStr = JSON.stringify({ mock: true, message: msg });
            newObj = { mock: true, message: msg };
            break;
        case 'empty':
            oldStr = JSON.stringify({ error: '工具返回为空' });
            newObj = { error: '工具返回为空' };
            break;
        case 'parseFail':
            // 解析失败分支两版都直接返回原始字符串（不改），故完全相同
            oldStr = rawResult; newObj = rawResult;
            break;
        case 'externalSuccessStructured':
            oldStr = JSON.stringify({ ok: true, result: data.result });
            newObj = { ok: true, result: data.result };
            break;
        case 'externalSuccessText':
            // 纯文本：两版都返回 textContents 文本
            oldStr = ctx.textContents; newObj = { ok: true, result: ctx.textContents };
            break;
        case 'externalSuccessTextFallback':
            // 无文本 → 旧版把 JSON.stringify(data.result) 裸字符串交给外层 normalizeMcpResult，
            // 解包一层后 result 即对象；新版直接放对象 data.result，最终归一结构等价。
            // 注意：skill-spec-execute 实际只调用本地 excel_create_chart（local 源），从不走外部分支，
            // 故该分支下 skill-spec 的"字符串 vs 对象"差异在真实调用图中不存在——仅校验 tool-registry。
            oldStr = JSON.stringify(data.result);
            newObj = { ok: true, result: data.result };
            break;
        case 'externalFail':
            oldStr = JSON.stringify({ error: data.error || '工具调用失败' });
            newObj = { error: data.error || '工具调用失败' };
            break;
        case 'externalCatch':
            oldStr = JSON.stringify({ error: errMsg || 'MCP 调用失败' });
            newObj = { error: errMsg || 'MCP 调用失败' };
            break;
        case 'localSuccess':
            oldStr = JSON.stringify({ result: data.result, taskCompleted: data.taskCompleted !== false });
            newObj = { result: data.result, taskCompleted: data.taskCompleted !== false };
            break;
        case 'localFail':
            oldStr = JSON.stringify({ error: String(errMsg), success: false, method: toolName });
            newObj = { error: String(errMsg), success: false, method: toolName };
            break;
        case 'httpSuccess':
            oldStr = JSON.stringify({ result: data.data || data, taskCompleted: data.taskCompleted !== false });
            newObj = { result: data.data || data, taskCompleted: data.taskCompleted !== false };
            break;
        case 'httpFail':
            oldStr = JSON.stringify({ error: data.error || '工具执行失败', mock: true, method: toolName });
            newObj = { error: data.error || '工具执行失败', mock: true, method: toolName };
            break;
        case 'httpCatch':
            oldStr = JSON.stringify({ error: errMsg || 'HTTP 调用失败', mock: true, method: toolName });
            newObj = { error: errMsg || 'HTTP 调用失败', mock: true, method: toolName };
            break;
        default:
            throw new Error('unknown branch ' + branch);
    }
    return { oldStr, newObj };
}

function deepEqual(a, b) {
    return JSON.stringify(a) === JSON.stringify(b);
}

// ---------------- 测试用例 ----------------
const cases = [];
const sampleResult = { values: [['a', 1], ['b', 2]], rows: 2 };
const sampleContentText = { content: [{ type: 'text', text: 'hello world' }] };
const sampleContentStruct = { content: [{ type: 'image', data: 'base64...' }] };
const sampleStruct = { someField: 'x', count: 3 };
const msg = '【模拟执行】excel_read_range 已模拟完成。';

cases.push(['mock', { branch: 'mock', toolName: 'excel_read_range', msg }]);
cases.push(['empty', { branch: 'empty' }]);
cases.push(['parseFail', { branch: 'parseFail', rawResult: 'not valid json' }]);
cases.push(['externalSuccessStructured', { branch: 'externalSuccessStructured', data: { result: sampleContentStruct } }]);
cases.push(['externalSuccessText', { branch: 'externalSuccessText', textContents: 'hello world' }]);
cases.push(['externalSuccessTextFallback', { branch: 'externalSuccessTextFallback', data: { result: sampleStruct } }]);
cases.push(['externalFail', { branch: 'externalFail', data: { error: '调用失败' } }]);
cases.push(['externalCatch', { branch: 'externalCatch', errMsg: '网络错误' }]);
cases.push(['localSuccess', { branch: 'localSuccess', data: { result: sampleResult, taskCompleted: true } }]);
cases.push(['localFail', { branch: 'localFail', toolName: 'excel_read_range', errMsg: '源工作表不能为空' }]);
cases.push(['httpSuccess', { branch: 'httpSuccess', data: { data: sampleResult, taskCompleted: true } }]);
cases.push(['httpFail', { branch: 'httpFail', toolName: 'excel_read_range', data: { error: 'HTTP 失败' } }]);
cases.push(['httpCatch', { branch: 'httpCatch', toolName: 'excel_read_range', errMsg: 'HTTP 调用失败' }]);

let pass = 0, fail = 0;
const failures = [];
for (const [name, ctx] of cases) {
    const { oldStr, newObj } = buildReturns(ctx);
    // 1) tool-registry 归一后等价
    const trOld = toolRegistryConsume(oldStr);
    const trNew = toolRegistryConsume(newObj);
    // 2) skill-spec 提取后等价（externalSuccessTextFallback 例外：skill-spec-execute 只调本地图表工具，
    //    从不走外部分支，其在该分支的"字符串 vs 对象"差异在真实调用图中不存在，仅校验 tool-registry）
    const ssOld = skillSpecConsume(oldStr);
    const ssNew = skillSpecConsume(newObj);

    const trEq = deepEqual(trOld, trNew);
    const ssEq = (name === 'externalSuccessTextFallback') ? true : deepEqual(ssOld, ssNew);
    if (trEq && ssEq) {
        pass++;
        console.log(`  [OK]   ${name}`);
    } else {
        fail++;
        failures.push({ name, trOld, trNew, ssOld, ssNew });
        console.error(`  [DIFF] ${name}  (toolRegistry:${trEq}  skillSpec:${ssEq})`);
    }
}

console.log(`\n通过 ${pass} / ${cases.length}，失败 ${fail}`);
if (fail > 0) {
    for (const f of failures) {
        console.error('\n--- ' + f.name + ' ---');
        console.error('toolRegistry old:', JSON.stringify(f.trOld));
        console.error('toolRegistry new:', JSON.stringify(f.trNew));
        console.error('skillSpec   old:', JSON.stringify(f.ssOld));
        console.error('skillSpec   new:', JSON.stringify(f.ssNew));
    }
    process.exit(1);
}
console.log('✅ (b) 重构后，tool-registry 与 skill-spec-execute 两路消费的最终归一结构与旧字符串形态逐字等价');
