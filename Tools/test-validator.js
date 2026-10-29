const validator = require('./vba-validator');

const testCode = `Sub BatchRenameSheets()
  Dim ws As Worksheet
  Dim sheetName As String
  For Each ws In ThisWorkbook.Worksheets
    sheetName = ws.Range("A1").Value
    On Error Resume Next
    ws.Name = sheetName
    On Error GoTo 0
  Next ws
End Sub`;

console.log('=== VBA 验证器测试 ===');
console.log('测试代码（您遇到的问题场景）:\n');
console.log(testCode);
console.log('\n');

const result = validator.validate(testCode);

console.log('✅ 验证通过:', result.isValid);
console.log('⚠️  风险等级:', result.riskLevel.message);
console.log('📊 风险评分:', result.score + '/4');
console.log('🔍 发现问题:', result.issues.length, '个');
console.log('💡 改进建议:', result.suggestions.length, '个');
console.log('⚡ 警告信息:', result.warnings.length, '个');
console.log('🔧 代码已修复:', result.fixedCode !== testCode ? '是' : '否');

if (result.issues.length > 0) {
  console.log('\n--- 发现的问题 ---');
  result.issues.forEach((issue, idx) => {
    console.log((idx+1) + '. [' + issue.riskColor + '] ' + issue.type);
    console.log('   建议: ' + issue.suggestion);
  });
}

if (result.warnings.length > 0) {
  console.log('\n--- 详细警告 ---');
  result.warnings.forEach((warning, idx) => {
    console.log((idx+1) + '. [' + warning.severity.toUpperCase() + '] ' + warning.type);
    console.log('   ' + warning.message);
    console.log('   解决: ' + warning.solution);
  });
}

if (result.fixedCode !== testCode) {
  console.log('\n--- 修复后的代码预览（前800字符） ---');
  console.log(result.fixedCode.substring(0, 800));
  console.log('\n... (完整代码已保存到 fixedCode 字段)');
}

console.log('\n=== 测试完成 ===');