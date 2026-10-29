const fs = require('fs');
const path = require('path');

class VBAValidator {
    constructor() {
        this.riskLevels = {
            LOW: { score: 1, color: '🟢', message: '低风险 - 可直接执行' },
            MEDIUM: { score: 2, color: '🟡', message: '中风险 - 建议审查后执行' },
            HIGH: { score: 3, color: '🟠', message: '高风险 - 必须确认后执行' },
            CRITICAL: { score: 4, color: '🔴', message: '严重风险 - 强烈建议修改' }
        };

        this.dangerousPatterns = [
            { pattern: /Kill\s+/i, risk: 'CRITICAL', type: '文件删除', fix: '请使用更安全的文件操作方法，或添加确认对话框' },
            { pattern: /\.Delete\b/i, risk: 'HIGH', type: '对象删除', fix: '添加 If MsgBox("确认删除?", vbYesNo) = vbYes Then 判断' },
            { pattern: /Shell\s*\(/i, risk: 'HIGH', type: '系统命令', fix: '避免使用 Shell 执行外部命令' },
            { pattern: /SendKeys/i, risk: 'MEDIUM', type: '键盘模拟', fix: 'SendKeys 不稳定，建议使用其他自动化方法' },
            { pattern: /On Error Resume Next(?!\s*'\s*临时)/i, risk: 'MEDIUM', type: '错误屏蔽', fix: '错误被静默忽略，建议添加错误处理逻辑' },
            { pattern: /\.Name\s*=/i, risk: 'MEDIUM', type: '重命名操作', fix: '需要验证名称合法性（长度、字符、重名等）' },
            { pattern: /ThisWorkbook\.Save/i, risk: 'MEDIUM', type: '保存操作', fix: '建议先备份或添加确认提示' },
            { pattern: /Application\.DisplayAlerts\s*=\s*False/i, risk: 'MEDIUM', type: '禁用警告', fix: '确保在操作后恢复 DisplayAlerts = True' },
            { pattern: /Workbooks\.Open/i, risk: 'LOW', type: '打开文件', fix: '确保文件路径存在且可访问' },
            { pattern: /Workbooks\.Add/i, risk: 'LOW', type: '新建工作簿', fix: '无特殊风险' }
        ];

        this.requiredPatterns = [
            { pattern: /Option Explicit/i, type: '变量声明', importance: '高', suggestion: '强制声明所有变量，避免拼写错误' },
            { pattern: /On Error GoTo|On Error Resume Next.*\n.*Err\.Number/i, type: '错误处理', importance: '高', suggestion: '添加结构化错误处理' },
            { pattern: /Dim\s+.*As\s+/i, type: '变量类型', importance: '中', suggestion: '明确指定变量类型提高性能' }
        ];
    }

    validate(code) {
        const result = {
            isValid: true,
            riskLevel: this.riskLevels.LOW,
            issues: [],
            suggestions: [],
            warnings: [],
            fixedCode: code,
            score: 0,
            executionPlan: null
        };

        console.log('\n[VBA Validator] 开始验证代码...');
        console.log('[VBA Validator] 原始代码长度:', code.length);

        let maxRiskScore = 0;
        
        for (const danger of this.dangerousPatterns) {
            if (danger.pattern.test(code)) {
                const riskInfo = this.riskLevels[danger.risk];
                result.issues.push({
                    type: danger.type,
                    risk: danger.risk,
                    riskColor: riskInfo.color,
                    message: `⚠️ ${danger.type}: ${riskInfo.message}`,
                    suggestion: danger.fix,
                    pattern: danger.pattern.toString()
                });

                if (riskInfo.score > maxRiskScore) {
                    maxRiskScore = riskInfo.score;
                    result.riskLevel = riskInfo;
                }

                if (danger.risk === 'CRITICAL') {
                    result.isValid = false;
                }
            }
        }

        result.score = maxRiskScore;

        for (const req of this.requiredPatterns) {
            if (!req.pattern.test(code)) {
                result.suggestions.push({
                    type: req.type,
                    importance: req.importance,
                    message: `💡 建议: 添加${req.type}`,
                    detail: req.suggestion
                });
            }
        }

        const specificIssues = this.checkSpecificPatterns(code);
        result.warnings = specificIssues;

        if (result.issues.length > 0 || result.suggestions.length > 0) {
            result.fixedCode = this.autoFix(code, result);
        }

        result.executionPlan = this.generateExecutionPlan(result);

        this.logValidationResult(result);

        return result;
    }

    checkSpecificPatterns(code) {
        const warnings = [];

        if (code.includes('ws.Name = ') || code.includes('.Name =')) {
            if (!code.includes('CleanWorksheetName') && !code.includes('ValidateSheetName')) {
                warnings.push({
                    type: '工作表命名',
                    severity: 'high',
                    message: '❌ 缺少名称验证函数 - 可能导致运行时错误1004',
                    solution: '需要添加 CleanWorksheetName() 和 WorksheetExists() 辅助函数'
                });
            }
        }

        if ((code.match(/For Each/g) || []).length > 0) {
            if (!code.includes('On Error') && !code.includes('Error Handling')) {
                warnings.push({
                    type: '循环操作',
                    severity: 'medium',
                    message: '⚠️ 循环中没有错误处理 - 单个失败可能导致整个循环中断',
                    solution: '在循环体内添加 On Error Resume Next 和错误计数器'
                });
            }
        }

        if (code.includes('Range(') && !code.includes('IsEmpty') && !code.includes('IsError')) {
            warnings.push({
                type: '单元格读取',
                severity: 'medium',
                message: '⚠️ 直接读取单元格值未做空值检查',
                solution: '添加 If Not IsEmpty(Range(...).Value) Then 检查'
            });
        }

        const subCount = (code.match(/Sub\s+\w+/g) || []).length;
        if (subCount === 0 && code.length > 50) {
            warnings.push({
                type: '代码结构',
                severity: 'low',
                message: 'ℹ️ 代码未封装在 Sub/Function 中',
                solution: '建议将代码封装为 Sub 或 Function 提高可维护性'
            });
        }

        return warnings;
    }

    autoFix(code, validationResult) {
        let fixedCode = code;

        if (!fixedCode.includes('Option Explicit')) {
            fixedCode = 'Option Explicit\n\n' + fixedCode;
        }

        const hasRenameOperation = fixedCode.includes('.Name =');
        const needsNameValidation = hasRenameOperation && 
                                   !fixedCode.includes('CleanWorksheetName') &&
                                   !fixedCode.includes('ValidateSheetName');

        if (needsNameValidation) {
            const helperFunctions = this.generateHelperFunctions();
            
            if (!fixedCode.endsWith('\n')) {
                fixedCode += '\n';
            }
            
            fixedCode += '\n' + '='.repeat(80) + '\n';
            fixedCode += '= 自动生成的辅助函数（VBA Validator 注入）\n';
            fixedCode += '='.repeat(80) + '\n\n';
            fixedCode += helperFunctions;
        }

        if (validationResult.issues.some(i => i.type === '错误屏蔽')) {
            fixedCode = this.addStructuredErrorHandling(fixedCode);
        }

        if (hasRenameOperation) {
            fixedCode = this.injectNameValidation(fixedCode);
        }

        return fixedCode;
    }

    generateHelperFunctions() {
        return `' ============================================================
' 【自动生成】VBA 安全辅助函数
' 由 SmartPilot VBA Validator 自动注入
' 生成时间: ${new Date().toISOString()}
' ============================================================

' 清理工作表名称中的非法字符
Private Function CleanWorksheetName(ByVal nameStr As String) As String
    Dim illegalChars As Variant
    Dim i As Integer
    
    nameStr = Trim(nameStr)
    
    If Len(nameStr) = 0 Then
        CleanWorksheetName = ""
        Exit Function
    End If
    
    If Len(nameStr) > 31 Then
        nameStr = Left(nameStr, 31)
    End If
    
    illegalChars = Array("/", "\", "[", "]", ":", "*", "?", "|")
    
    For i = LBound(illegalChars) To UBound(illegalChars)
        nameStr = Replace(nameStr, illegalChars(i), "_")
    Next i
    
    nameStr = Replace(nameStr, Chr(0), "")
    CleanWorksheetName = nameStr
End Function

' 检查工作表是否已存在
Private Function WorksheetExists(ByVal wsName As String, Optional ByVal wb As Workbook = Nothing) As Boolean
    Dim ws As Worksheet
    Dim targetWb As Workbook
    
    On Error Resume Next
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If
    
    Set ws = targetWb.Worksheets(wsName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' 安全重命名工作表（带完整验证）
Private Function SafeRenameWorksheet(ByVal ws As Worksheet, ByVal newName As String, _
                                     ByRef errorMsg As String) As Boolean
    Dim cleanName As String
    
    On Error GoTo ErrorHandler
    
    cleanName = CleanWorksheetName(newName)
    
    If Len(cleanName) = 0 Then
        errorMsg = "工作表名称不能为空"
        SafeRenameWorksheet = False
        Exit Function
    End If
    
    If cleanName = ws.Name Then
        errorMsg = ""
        SafeRenameWorksheet = True
        Exit Function
    End If
    
    If WorksheetExists(cleanName) Then
        errorMsg = "工作表 '" & cleanName & "' 已存在"
        SafeRenameWorksheet = False
        Exit Function
    End If
    
    ws.Name = cleanName
    errorMsg = ""
    SafeRenameWorksheet = True
    Exit Function
    
ErrorHandler:
    errorMsg = "重命名失败: " & Err.Description
    SafeRenameWorksheet = False
End Function

' 获取安全的单元格文本值
Private Function GetSafeCellText(ByVal rng As Range, Optional ByVal defaultValue As String = "") As String
    On Error Resume Next
    If rng Is Nothing Then
        GetSafeCellText = defaultValue
        Exit Function
    End If
    
    If IsEmpty(rng.Value) Or IsError(rng.Value) Or IsNull(rng.Value) Then
        GetSafeCellText = defaultValue
    Else
        GetSafeCellText = CStr(Trim(rng.Value))
    End If
    On Error GoTo 0
End Function
`;
    }

    addStructuredErrorHandling(code) {
        const lines = code.split('\n');
        const processedLines = [];
        let inSub = false;
        let addedErrorHandler = false;

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            processedLines.push(line);

            if (/^\s*Sub\s+\w+/i.test(line)) {
                inSub = true;
                addedErrorHandler = false;
            }

            if (inSub && /^\s*End Sub/i.test(line) && !addedErrorHandler) {
                const indent = line.match(/^\s*/)[0];
                processedLines.splice(-1, 0, 
                    `${indent}Exit Sub\n`,
                    `${indent}ErrorHandler:\n`,
                    `${indent}    Debug.Print "错误 [${line.match(/Sub\s+(\w+)/i)?.[1] || 'Unknown'}]: " & Err.Number & " - " & Err.Description\n`,
                    `${indent}    ' 可在此处添加错误恢复逻辑`
                );
                inSub = false;
                addedErrorHandler = true;
            }
        }

        return processedLines.join('\n');
    }

    injectNameValidation(code) {
        let modifiedCode = code;

        modifiedCode = modifiedCode.replace(
            /(\w+)\.Name\s*=\s*(\w+(?:\([^)]*\))?)/gi,
            (match, obj, value) => {
                return `If Not SafeRenameWorksheet(${obj}, ${value}, errMsg) Then
            Debug.Print "⚠️ 重命名失败: " & errMsg
        Else
            ${obj}.Name = CleanWorksheetName(${value})
        End If`;
            }
        );

        return modifiedCode;
    }

    generateExecutionPlan(validationResult) {
        const plan = {
            canExecute: validationResult.isValid && validationResult.score <= 2,
            requiresConfirmation: validationResult.score >= 2,
            steps: [],
            estimatedRisk: validationResult.riskLevel.message
        };

        if (validationResult.issues.length > 0) {
            plan.steps.push({
                step: 1,
                action: 'review_issues',
                description: '审查并确认所有风险项',
                items: validationResult.issues.map(i => ({
                    issue: i.type,
                    risk: i.risk,
                    suggestion: i.suggestion
                }))
            });
        }

        if (validationResult.fixedCode !== validationResult.code?.originalCode) {
            plan.steps.push({
                step: plan.steps.length + 1,
                action: 'use_fixed_code',
                description: '使用已修复的代码版本（已注入安全检查）',
                changes: validationResult.suggestions.length
            });
        }

        plan.steps.push({
            step: plan.steps.length + 1,
            action: 'execute',
            description: validationResult.canExecute ? '执行代码' : '需要人工确认后执行',
            mode: validationResult.score <= 1 ? 'auto' : 'manual'
        });

        return plan;
    }

    logValidationResult(result) {
        console.log('\n' + '='.repeat(60));
        console.log('[VBA Validator] 验证结果报告');
        console.log('='.repeat(60));
        console.log(`✅ 代码有效性: ${result.isValid ? '有效' : '无效'}`);
        console.log(`⚠️  风险等级: ${result.riskLevel.color} ${result.riskLevel.message}`);
        console.log(`📊 风险评分: ${result.score}/4`);
        console.log(`🔍 发现问题: ${result.issues.length} 个`);
        console.log(`💡 改进建议: ${result.suggestions.length} 个`);
        console.log(`⚡ 警告信息: ${result.warnings.length} 个`);

        if (result.issues.length > 0) {
            console.log('\n--- 问题详情 ---');
            result.issues.forEach((issue, idx) => {
                console.log(`${idx + 1}. [${issue.riskColor}] ${issue.type}`);
                console.log(`   建议: ${issue.suggestion}`);
            });
        }

        if (result.warnings.length > 0) {
            console.log('\n--- 警告详情 ---');
            result.warnings.forEach((warning, idx) => {
                console.log(`${idx + 1}. [${warning.severity.toUpperCase()}] ${warning.type}`);
                console.log(`   ${warning.message}`);
                console.log(`   解决: ${warning.solution}`);
            });
        }

        console.log('\n--- 执行计划 ---');
        console.log(`可以执行: ${result.executionPlan.canExecute ? '是' : '否'}`);
        console.log(`需要确认: ${result.executionPlan.requiresConfirmation ? '是' : '否'}`);
        result.executionPlan.steps.forEach(step => {
            console.log(`步骤 ${step.step}: ${step.description}`);
        });

        console.log('='.repeat(60) + '\n');
    }

    generateReport(validationResult) {
        const report = {
            timestamp: new Date().toISOString(),
            summary: {
                isValid: validationResult.isValid,
                riskLevel: validationResult.riskLevel.message,
                riskScore: validationResult.score,
                issuesCount: validationResult.issues.length,
                suggestionsCount: validationResult.suggestions.length
            },
            details: {
                issues: validationResult.issues,
                suggestions: validationResult.suggestions,
                warnings: validationResult.warnings
            },
            executionPlan: validationResult.executionPlan,
            codeStats: {
                originalLength: validationResult.code?.length || 0,
                fixedLength: validationResult.fixedCode.length,
                linesAdded: validationResult.fixedCode.split('\n').length - (validationResult.code?.split('\n').length || 0)
            }
        };

        return report;
    }
}

module.exports = new VBAValidator();