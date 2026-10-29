Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic

Namespace Composition

    ''' <summary>
    ''' 服务生命周期枚举
    ''' </summary>
    Public Enum ServiceLifetime
        Singleton
        Transient
    End Enum

    ''' <summary>
    ''' 轻量级服务注册描述符（替代 Microsoft.Extensions.DependencyInjection.ServiceDescriptor）
    ''' </summary>
    Public Class ServiceDescriptor
        Public Property ServiceType As Type
        Public Property ImplementationType As Type
        Public Property ImplementationFactory As Func(Of Object)
        Public Property Lifetime As ServiceLifetime
    End Class

    ''' <summary>
    ''' 服务注册表（替代 IServiceCollection）
    ''' </summary>
    Public Class ServiceCollection
        Inherits List(Of ServiceDescriptor)
    End Class

    ''' <summary>
    ''' 应用服务定位器 / Composition Root
    '''
    ''' 【渐进式 DI 设计】
    ''' ------------------------------------------------------------------
    ''' Phase 1（当前）：双轨运行
    '''   • 现有代码仍然可以使用 Class.Instance / GetInstance() 访问单例
    '''   • Class.Instance 的 getter 在内部调用 AppServiceProvider.Get(Of T)()
    '''   • 服务在 ConfigureServices() 中注册生命周期（Singleton / Transient）
    '''
    ''' Phase 2（后续）：完全迁移
    '''   • 构造函数注入替代 Class.Instance
    '''   • 移除所有 Shared Instance 属性
    '''   • 所有服务通过 DI 容器解析，便于单元测试 Mock
    ''' ------------------------------------------------------------------
    '''
    ''' 【关闭流程简化】
    '''   原来 ThisAddIn.Dispose 需要按顺序手动释放 10+ 单例；
    '''   现在只需 AppServiceProvider.Instance.Dispose()，
    '''   容器会自动释放所有实现了 IDisposable 的 Singleton 服务。
    ''' </summary>
    Public NotInheritable Class AppServiceProvider
        Implements IDisposable

        Private Shared ReadOnly _bootstrapLock As New Object()
        Private Shared _instance As AppServiceProvider
        Private Shared _isConfigured As Boolean = False

        Private ReadOnly _serviceCollection As ServiceCollection
        Private ReadOnly _singletonInstances As New Dictionary(Of Type, Object)()
        Private ReadOnly _disposables As New List(Of IDisposable)()
        Private _disposedValue As Boolean = False
        Private _built As Boolean = False

        ''' <summary>
        ''' 全局单例入口（启动引导器本身使用经典单例模式）
        ''' </summary>
        Public Shared ReadOnly Property Instance As AppServiceProvider
            Get
                If _instance Is Nothing Then
                    SyncLock _bootstrapLock
                        If _instance Is Nothing Then
                            _instance = New AppServiceProvider()
                        End If
                    End SyncLock
                End If
                Return _instance
            End Get
        End Property

        ''' <summary>
        ''' 是否已经完成 ConfigureServices()
        ''' </summary>
        Public Shared ReadOnly Property IsConfigured As Boolean
            Get
                Return _isConfigured
            End Get
        End Property

        ''' <summary>
        ''' 构造函数 — 创建 ServiceCollection，但不 Build（延迟到 ConfigureServices）
        ''' </summary>
        Private Sub New()
            _serviceCollection = New ServiceCollection()
        End Sub

        ''' <summary>
        ''' 注册默认服务清单 — 由 ThisAddIn 在初始化早期调用一次
        ''' 传入的 extendAction 允许 ThisAddIn 追加注册 Excel/UI 层的服务
        ''' </summary>
        Public Sub ConfigureServices(Optional extendAction As Action(Of ServiceCollection) = Nothing)
            SyncLock _bootstrapLock
                If _isConfigured Then
                    Throw New InvalidOperationException("AppServiceProvider.ConfigureServices() 只能调用一次。")
                End If

                RegisterCoreServices(_serviceCollection)
                If extendAction IsNot Nothing Then
                    extendAction(_serviceCollection)
                End If

                _built = True
                _isConfigured = True
            End SyncLock
        End Sub

        ''' <summary>
        ''' 获取服务。未 ConfigureServices 时会抛 InvalidOperationException。
        ''' </summary>
        Public Function [Get](Of T)() As T
            EnsureConfigured()
            Dim result As Object = GetService(GetType(T), True)
            If result Is Nothing Then
                Throw New InvalidOperationException($"未注册服务: {GetType(T).FullName}")
            End If
            Return CType(result, T)
        End Function

        ''' <summary>
        ''' 尝试获取服务（不抛异常），未注册返回 Nothing
        ''' </summary>
        Public Function TryGet(Of T As Class)() As T
            If Not _isConfigured Then Return Nothing
            Dim result As Object = GetService(GetType(T), False)
            If result Is Nothing Then Return Nothing
            Return CType(result, T)
        End Function

        ''' <summary>
        ''' 返回自身（兼容接口，满足原有 GetServiceProvider() 签名场景）
        ''' </summary>
        Public Function GetServiceProvider() As Object
            EnsureConfigured()
            Return Me
        End Function

        ''' <summary>
        ''' 测试场景专用：重置容器（非测试代码不要调用）
        ''' </summary>
        Public Shared Sub ResetForTesting()
            If _instance IsNot Nothing Then
                _instance.Dispose()
                _instance = Nothing
            End If
            _isConfigured = False
        End Sub

        ''' <summary>
        ''' 释放容器 — 触发所有 Singleton IDisposable 的 Dispose()
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Private Sub Dispose(disposing As Boolean)
            If Not _disposedValue Then
                If disposing Then
                    ' 逆序释放，先注册的后释放
                    For i As Integer = _disposables.Count - 1 To 0 Step -1
                        Try
                            _disposables(i).Dispose()
                        Catch ex As Exception  ' 
                        End Try
                    Next
                    _disposables.Clear()
                    _singletonInstances.Clear()
                End If
                _disposedValue = True
            End If
        End Sub

        ''' <summary>
        ''' 核心解析逻辑
        ''' </summary>
        Private Function GetService(serviceType As Type, throwOnError As Boolean) As Object
            SyncLock _bootstrapLock
                ' 查找匹配的描述符
                Dim descriptor As ServiceDescriptor = Nothing
                For Each d As ServiceDescriptor In _serviceCollection
                    If d.ServiceType Is serviceType Then
                        descriptor = d
                        Exit For
                    End If
                Next

                If descriptor Is Nothing Then
                    If throwOnError Then
                        Throw New InvalidOperationException($"服务未注册: {serviceType.FullName}")
                    End If
                    Return Nothing
                End If

                Select Case descriptor.Lifetime
                    Case ServiceLifetime.Singleton
                        If _singletonInstances.ContainsKey(serviceType) Then
                            Return _singletonInstances(serviceType)
                        End If
                        Dim instance As Object = CreateInstance(descriptor)
                        _singletonInstances(serviceType) = instance
                        ' 跟踪可释放对象
                        If TypeOf instance Is IDisposable Then
                            _disposables.Add(CType(instance, IDisposable))
                        End If
                        Return instance

                    Case ServiceLifetime.Transient
                        Return CreateInstance(descriptor)

                    Case Else
                        Return Nothing
                End Select
            End SyncLock
        End Function

        ''' <summary>
        ''' 根据描述符创建实例
        ''' </summary>
        Private Shared Function CreateInstance(descriptor As ServiceDescriptor) As Object
            If descriptor.ImplementationFactory IsNot Nothing Then
                Return descriptor.ImplementationFactory()
            End If

            Dim implType As Type = If(descriptor.ImplementationType, descriptor.ServiceType)

            ' 优先尝试无参构造函数（本项目所有服务都是无参单例）
            Dim ctor As Reflection.ConstructorInfo = implType.GetConstructor(Type.EmptyTypes)
            If ctor IsNot Nothing Then
                Return ctor.Invoke(New Object() {})
            End If

            ' 尝试找参数最少的构造函数（简单激活器）
            Dim ctors As Reflection.ConstructorInfo() = implType.GetConstructors()
            If ctors.Length > 0 Then
                Array.Sort(ctors, Function(a, b) a.GetParameters().Length.CompareTo(b.GetParameters().Length))
                Dim minCtor As Reflection.ConstructorInfo = ctors(0)
                Dim params As Reflection.ParameterInfo() = minCtor.GetParameters()
                Dim args As Object() = New Object(params.Length - 1) {}
                For i As Integer = 0 To params.Length - 1
                    args(i) = Nothing
                Next
                Return minCtor.Invoke(args)
            End If

            Throw New InvalidOperationException($"无法创建实例: {implType.FullName}")
        End Function

        ''' <summary>
        ''' 核心服务注册（不含 Excel COM 依赖的纯逻辑/配置服务）
        ''' 所有服务在这里集中声明，便于新开发者快速建立服务全景图
        ''' </summary>
        Private Shared Sub RegisterCoreServices(services As ServiceCollection)
            ' —— 日志基础设施（按注册顺序从细到粗）——
            services.AddSingleton(Of Utils.AppLogger)()
            services.AddSingleton(Of Utils.ChatLogger)()
            ' 日志组件经 ILogService 统一抽象，由 LogServiceManager 提供多目标输出支持
            services.AddSingleton(Of Utils.ILogService, Utils.LogServiceManager)()

            ' —— 缓存层 ——
            services.AddSingleton(Of Utils.RequestCache)()

            ' —— 配置管理 ——
            services.AddSingleton(Of Utils.ConfigurationManager)()
            services.AddSingleton(Of Services.Mcp.McpConfigManager)()

            ' —— 策略 / 工具类 ——
            services.AddSingleton(Of Utils.RetryPolicy)()
            services.AddSingleton(Of Services.OcrFormatFormatter)()
            services.AddSingleton(Of Skills.SkillApprovalGate)(Function() Skills.SkillApprovalGate.Instance)
            services.AddSingleton(Of Utils.ConfigFileWatcher)()

            ' —— AI 辅助类（纯逻辑，无 Excel COM 调用）——
            services.AddSingleton(Of Services.PromptGenerator)()

            ' —— 容错恢复 ——
            services.AddSingleton(Of Services.ErrorRecovery)()
        End Sub

        Private Sub EnsureConfigured()
            If Not _isConfigured Then
                Throw New InvalidOperationException(
                    "AppServiceProvider 尚未配置。请在 ThisAddIn 初始化时调用 ConfigureServices()。" &
                    "如需单元测试，请先手动调用 AppServiceProvider.Instance.ConfigureServices()。")
            End If
        End Sub

    End Class

    ''' <summary>
    ''' 扩展方法：与 Microsoft.Extensions.DependencyInjection API 风格保持一致，方便后续切换
    ''' </summary>
    Public Module ServiceCollectionExtensions

        ''' <summary>
        ''' 注册 Singleton（实现类型即服务类型）
        ''' </summary>
        <System.Runtime.CompilerServices.Extension()>
        Public Function AddSingleton(Of TService)(services As ServiceCollection) As ServiceCollection
            services.Add(New ServiceDescriptor With {
                .ServiceType = GetType(TService),
                .ImplementationType = GetType(TService),
                .Lifetime = ServiceLifetime.Singleton
            })
            Return services
        End Function

        ''' <summary>
        ''' 注册 Singleton（服务类型 + 实现类型）
        ''' </summary>
        <System.Runtime.CompilerServices.Extension()>
        Public Function AddSingleton(Of TService, TImplementation As TService)(services As ServiceCollection) As ServiceCollection
            services.Add(New ServiceDescriptor With {
                .ServiceType = GetType(TService),
                .ImplementationType = GetType(TImplementation),
                .Lifetime = ServiceLifetime.Singleton
            })
            Return services
        End Function

        ''' <summary>
        ''' 注册 Singleton（使用工厂方法）
        ''' </summary>
        <System.Runtime.CompilerServices.Extension()>
        Public Function AddSingleton(Of TService)(services As ServiceCollection, factory As Func(Of TService)) As ServiceCollection
            services.Add(New ServiceDescriptor With {
                .ServiceType = GetType(TService),
                .ImplementationFactory = Function() factory(),
                .Lifetime = ServiceLifetime.Singleton
            })
            Return services
        End Function

        ''' <summary>
        ''' 注册 Transient（实现类型即服务类型）
        ''' </summary>
        <System.Runtime.CompilerServices.Extension()>
        Public Function AddTransient(Of TService)(services As ServiceCollection) As ServiceCollection
            services.Add(New ServiceDescriptor With {
                .ServiceType = GetType(TService),
                .ImplementationType = GetType(TService),
                .Lifetime = ServiceLifetime.Transient
            })
            Return services
        End Function

    End Module

End Namespace