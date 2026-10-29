// VBA 模板安全静态校验（来自 Tools/vba-validator.js，阶段 4 P1 交付）
const { test } = require('node:test');
const { strict: assert } = require('node:assert');
const path = require('path');

const validator = require(path.join(__dirname, '..', 'Tools', 'vba-validator.js'));

// validate() 内部会打大量日志，测试时静默之
function run(code) {
  const origLog = console.log;
  console.log = () => {};
  try {
    return validator.validate(code);
  } finally {
    console.log = origLog;
  }
}

const SAFE_SUB = `Option Explicit
Sub SafeOperation()
    On Error GoTo ErrorHandler
    Dim r As Range
    Set r = Range("A1")
    If IsEmpty(r.Value) Then Exit Sub
    Exit Sub
ErrorHandler:
    Debug.Print "err: " & Err.Description
End Sub`;

test('安全模板：isValid=true 且风险评分<=2', () => {
  const r = run(SAFE_SUB);
  assert.strictEqual(r.isValid, true);
  assert.ok(r.score <= 2, `score 应为低风险，实际 ${r.score}`);
  assert.ok(r.issues.length === 0 || r.issues.every(i => i.risk !== 'CRITICAL'));
});

test('危险模式逐一被识别且风险等级正确', () => {
  const cases = [
    ['Kill "C:\\tmp\\x"', 'CRITICAL', '文件删除'],
    ['Cells(1,1).Delete', 'HIGH', '对象删除'],
    ['Shell("cmd /c dir", vbNormalFocus)', 'HIGH', '系统命令'],
    ['SendKeys "{ENTER}"', 'MEDIUM', '键盘模拟'],
    ['On Error Resume Next', 'MEDIUM', '错误屏蔽'],
    ['ws.Name = newName', 'MEDIUM', '重命名操作'],
    ['Application.DisplayAlerts = False', 'MEDIUM', '禁用警告'],
  ];
  for (const [snippet, risk, type] of cases) {
    const r = run(`Option Explicit\nSub T()\n  ${snippet}\nEnd Sub`);
    const hit = r.issues.find(i => i.type === type);
    assert.ok(hit, `应识别 ${type}（${snippet}）`);
    assert.strictEqual(hit.risk, risk, `${type} 风险应为 ${risk}，实际 ${hit.risk}`);
    assert.ok(typeof hit.suggestion === 'string' && hit.suggestion.length > 0, `${type} 应含修复建议`);
  }
});

test('CRITICAL 使 isValid=false 且 executionPlan.canExecute=false', () => {
  const r = run('Sub T()\n  Kill "C:\\tmp\\x"\nEnd Sub');
  assert.strictEqual(r.isValid, false);
  assert.ok(r.executionPlan, '应返回 executionPlan');
  assert.strictEqual(r.executionPlan.canExecute, false);
});

test('缺少 Option Explicit/Dim 时给出建议', () => {
  const r = run('Sub T()\n  x = 1\nEnd Sub');
  const types = r.suggestions.map(s => s.type);
  assert.ok(types.includes('变量声明'), '应建议 Option Explicit');
});

test('autoFix：为缺 Option Explicit 的代码注入声明', () => {
  const r = run('Sub T()\n  Dim x As Integer\n  x = 1\nEnd Sub');
  assert.ok(r.fixedCode.startsWith('Option Explicit'), 'fixedCode 应以 Option Explicit 开头');
});

test('重命名操作 autoFix 注入名称校验辅助函数', () => {
  const r = run('Sub T()\n  ws.Name = "新表名"\nEnd Sub');
  assert.ok(r.fixedCode.includes('CleanWorksheetName'), '应注入 CleanWorksheetName');
  assert.ok(r.fixedCode.includes('SafeRenameWorksheet'), '应注入 SafeRenameWorksheet');
});