' ============================================================
' PathHelperStub.vb — 仅供 VB 测试工程编译 PptDesignEngine.vb 使用
'
' PptDesignEngine.vb 仅在生产路径 AnalyzeReferenceTemplate 中用到
' PathHelper.ConfigDirectory（加载 Config/ppt_templates.json）。PptDesignEngine
' 的纯决策链（AnalyzeData / BuildDesign / DesignFromData / Recommend* / GenerateInsights
' / ApplyColumnHints）不触碰该属性。
'
' 因此这里提供一个最小桩，只兑现"能编译通过"所需的最小公共表面，让测试
' 直接编译那份真实的 PptDesignEngine.vb 生产代码（单一事实源），而不为测试
' 复制/改写任何决策逻辑、也不引入 ConfigurationManager/PythonBootstrapper 等
' 无关依赖。
' ============================================================
Namespace Utils

    Public Class PathHelper

        ''' <summary>仅需编译通过；本测试套件不触发布局注册表 IO</summary>
        Public Shared ReadOnly Property ConfigDirectory As String =
            System.AppDomain.CurrentDomain.BaseDirectory

    End Class

End Namespace