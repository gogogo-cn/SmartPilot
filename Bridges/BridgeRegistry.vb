Option Explicit On
Option Infer On

Imports System.Collections.Concurrent
Imports ExcelAIVSTOAddIn.Bridges

Namespace Bridges

    Public Class BridgeRegistry

        Private Shared ReadOnly _lazyInstance As New Lazy(Of BridgeRegistry)(
            Function() New BridgeRegistry(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Private ReadOnly _bridges As New ConcurrentDictionary(Of String, IBridge)
        ' 按接口类型索引已注册 Bridge，供依赖倒置按抽象解析（不改变既有按名字 API）
        Private ReadOnly _byInterface As New ConcurrentDictionary(Of Type, IBridge)

        Public Shared ReadOnly Property Instance As BridgeRegistry
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Friend Sub New()
        End Sub

        Public Function Register(bridge As IBridge) As BridgeRegistry
            If bridge Is Nothing Then Return Me
            _bridges(bridge.BridgeName) = bridge
            IndexByInterface(bridge)
            System.Diagnostics.Debug.WriteLine($"[BridgeRegistry] 注册: {bridge.BridgeName}")
            Return Me
        End Function

        Public Function RegisterLazy(Of T As {IBridge, New})(bridgeName As String) As BridgeRegistry
            _bridges(bridgeName) = DirectCast(New LazyBridge(Of T)(bridgeName), IBridge)
            System.Diagnostics.Debug.WriteLine($"[BridgeRegistry] 延迟注册: {bridgeName}")
            Return Me
        End Function

        ''' <summary>
        ''' 将 Bridge 按其实现的公共抽象接口建立索引（IBridge 与具体类型除外），
        ''' 使消费方可通过 Resolve(Of TInterface) 按接口拿到强类型引用，实现依赖倒置。
        ''' 后注册的同接口 Bridge 覆盖先注册者（保留最近一次）。
        ''' </summary>
        Private Sub IndexByInterface(bridge As IBridge)
            Try
                Dim concreteType = bridge.GetType()
                For Each iface As Type In concreteType.GetInterfaces()
                    If iface.Name.StartsWith("I") AndAlso iface IsNot GetType(IBridge) AndAlso iface.IsInterface Then
                        _byInterface(iface) = bridge
                    End If
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[BridgeRegistry] 索引接口失败 {bridge.GetType().Name}: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 按抽象接口解析 Bridge 强类型引用（依赖倒置入口）。
        ''' 仅解析已注册为非延迟的实例；若目标以延迟方式注册且尚未创建，返回 Nothing 交由调用方决定。
        ''' </summary>
        Public Function Resolve(Of TInterface As Class)() As TInterface
            Dim bridge As IBridge = Nothing
            If _byInterface.TryGetValue(GetType(TInterface), bridge) Then
                Return TryCast(bridge, TInterface)
            End If
            Return Nothing
        End Function

        Default Public ReadOnly Property Item(bridgeName As String) As IBridge
            Get
                Dim bridge As IBridge = Nothing
                _bridges.TryGetValue(bridgeName, bridge)
                Return bridge
            End Get
        End Property

        Public Function GetBridge(Of T As {IBridge, Class})(bridgeName As String) As T
            Dim bridge = Me(bridgeName)
            Return TryCast(bridge, T)
        End Function

        Public ReadOnly Property AllBridges As IEnumerable(Of IBridge)
            Get
                Return _bridges.Values
            End Get
        End Property

        Public ReadOnly Property InitializedBridges As IEnumerable(Of IBridge)
            Get
                Return _bridges.Values.Where(Function(b) TypeOf b IsNot ILazyBridge OrElse DirectCast(b, ILazyBridge).IsValueCreated)
            End Get
        End Property

        Public Sub ShutdownAll()
            System.Diagnostics.Debug.WriteLine($"[BridgeRegistry] ShutdownAll 开始，共 {_bridges.Count} 个注册项...")

            Dim disposedCount = 0
            ' 快照后迭代：释放某个 Bridge 时可能触发其他 Bridge 的延迟初始化并再次 Register（重入），
            ' 直接 For Each 遍历 ConcurrentDictionary 在高并发重入下枚举语义不保证稳定
            For Each kvp In _bridges.ToArray()
                Try
                    Dim bridge = kvp.Value
                    If TypeOf bridge Is ILazyBridge Then
                        Dim lazy = DirectCast(bridge, ILazyBridge)
                        If Not lazy.IsValueCreated Then
                            System.Diagnostics.Debug.WriteLine($"[BridgeRegistry] 跳过未初始化: {kvp.Key}")
                            Continue For
                        End If
                        bridge = lazy.Value
                    End If

                    System.Diagnostics.Debug.WriteLine($"[BridgeRegistry] 释放: {bridge.BridgeName}")
                    bridge.Dispose()
                    disposedCount += 1
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[BridgeRegistry] 释放 {kvp.Key} 失败: {ex.Message}")
                End Try
            Next

            _bridges.Clear()
            System.Diagnostics.Debug.WriteLine($"[BridgeRegistry] ShutdownAll 完成，共释放 {disposedCount} 个 Bridge")
        End Sub

        Private Interface ILazyBridge
            ReadOnly Property IsValueCreated As Boolean
            ReadOnly Property Value As IBridge
        End Interface

        Private Class LazyBridge(Of T As {IBridge, New})
            Implements ILazyBridge

            Private ReadOnly _inner As New Lazy(Of T)(Function() New T(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)
            Private ReadOnly _name As String

            Public Sub New(name As String)
                _name = name
            End Sub

            Public ReadOnly Property IsValueCreated As Boolean Implements ILazyBridge.IsValueCreated
                Get
                    Return _inner.IsValueCreated
                End Get
            End Property

            Public ReadOnly Property Value As IBridge Implements ILazyBridge.Value
                Get
                    Return _inner.Value
                End Get
            End Property

            Public ReadOnly Property BridgeName As String
                Get
                    Return _name
                End Get
            End Property
        End Class

    End Class

End Namespace