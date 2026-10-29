const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..', 'Skills');
// 默认全量扫描；命令行传文件名则只校验这些
const argNames = process.argv.slice(2);
const names = argNames.length ? argNames
  : fs.readdirSync(dir).filter(f => f.endsWith('.md')).map(f => f.slice(0, -3));

// 极简 frontmatter 解析：顶层 key: value + 列表项计数
function parseFM(text) {
  const m = text.replace(/\r\n/g, '\n').match(/^---\n([\s\S]*?)\n---\n/);
  if (!m) return null;
  const lines = m[1].split('\n');
  const out = {};
  let curList = null;
  for (const ln of lines) {
    // 先判"只有 key 没有值"→ 开启列表收集（必须早于带值分支）
    const bare = ln.match(/^([A-Za-z_][A-Za-z0-9_]*):[ \t]*$/);
    if (bare) { curList = bare[1]; out[bare[1]] = []; continue; }
    const km = ln.match(/^([A-Za-z_][A-Za-z0-9_]*):[ \t]*(.*)$/);
    if (km) { curList = null; out[km[1]] = km[2].trim(); continue; }
    const lm = ln.match(/^[ \t]*-[ \t]+(.*)$/);
    if (lm && curList !== null) { out[curList].push(lm[1].trim()); continue; }
  }
  return out;
}

let bad = 0, soft = 0;
// 与 Skills/Engine/SkillDefinition.vb 的 SkillProcessingMode 枚举保持一致
const VALID_MODES = new Set(['VBACode', 'MultimodalLLM', 'Hybrid', 'LocalOCR', 'PythonCode', 'SkillSpec', 'Skill']);
for (const n of names) {
  const p = path.join(dir, n + '.md');
  if (!fs.existsSync(p)) { console.log('MISSING', n); bad++; continue; }
  const fm = parseFM(fs.readFileSync(p, 'utf8'));
  if (!fm) { console.log('NO-FM', n); bad++; continue; }
  const problems = [];   // 硬错误：会导致技能加载失败或路由错误
  const warns = [];      // 可选字段缺失：不影响加载
  if (fm.name !== n) problems.push('name!=filename (' + fm.name + ')');
  if (!fm.hosts) problems.push('缺 hosts');
  if (fm.processingMode && !VALID_MODES.has(fm.processingMode)) {
    problems.push('processingMode="' + fm.processingMode + '" 不在 SkillProcessingMode 枚举内');
  }
  // 以下在本项目为可选项，缺失只提示
  if (!fm.category) warns.push('缺 category');
  if (!fm.dependencies) warns.push('缺 dependencies');
  if (!Array.isArray(fm.keywords) || fm.keywords.length === 0) warns.push('缺 keywords');
  if (!Array.isArray(fm.quickPrompts) || fm.quickPrompts.length === 0) warns.push('缺 quickPrompts');
  if (!Array.isArray(fm.examples) || fm.examples.length === 0) warns.push('缺 examples');
  // examples 里反例应指向的技能必须真实存在（写错名字会让路由失效）
  if (Array.isArray(fm.examples)) {
    for (const ex of fm.examples) {
      const rm = ex.match(/反例：应走([A-Za-z_][A-Za-z0-9_]*)/);
      if (rm && !fs.existsSync(path.join(dir, rm[1] + '.md'))) warns.push('反例指向不存在的技能 ' + rm[1]);
    }
  }
  const tag = problems.length ? 'FAIL' : (warns.length ? 'WARN' : 'OK  ');
  if (problems.length) bad++; else if (warns.length) soft++;
  console.log(tag, n.padEnd(24), 'hosts=' + String(fm.hosts).padEnd(8),
    'pipeline=' + String(fm.pipeline).padEnd(7), 'mode=' + String(fm.processingMode).padEnd(10),
    'kw=' + (Array.isArray(fm.keywords) ? fm.keywords.length : 0),
    'ex=' + (Array.isArray(fm.examples) ? fm.examples.length : 0),
    'qp=' + (Array.isArray(fm.quickPrompts) ? fm.quickPrompts.length : 0),
    problems.length ? '<< ' + problems.join('; ') : (warns.length ? '<< ' + warns.join('; ') : ''));
}
console.log('\n共 ' + names.length + ' 个技能：硬错误 ' + bad + ' 个' + (soft ? '，可选字段缺失 ' + soft + ' 个' : ''));
process.exit(bad ? 1 : 0);
