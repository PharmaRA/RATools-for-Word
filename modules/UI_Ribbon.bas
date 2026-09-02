Attribute VB_Name = "UI_Ribbon"
Option Explicit

' =============================================
' Ribbon 生命周期与事件初始化
' v0.8 重构阶段5：样式引擎、文档命令、宏注册表已拆分至
' Mod_StyleEngine / Mod_DocCommands / Mod_MacroRegistry。
' 本模块只负责 Ribbon 装载、选项卡激活与应用级事件挂接。
' 公共变量 gRibbonUI 供全工程访问 Ribbon 句柄。
' =============================================

Public gRibbonUI      As IRibbonUI     ' 缓存 Ribbon（Public 以便访问）
Private mAppEvents    As clsAppEvents  ' 事件监听器实例

' Ribbon 装载完成时回调（customUI14.xml onLoad）
Public Sub Onload(ribbon As IRibbonUI)
    Set gRibbonUI = ribbon
    InitEvents
    ActivateRAToolsTab
End Sub

Public Sub InitEvents()
    If mAppEvents Is Nothing Then
        Set mAppEvents = New clsAppEvents
    End If
End Sub

' 供外部调用的激活入口（使用延迟异步模式）
Public Sub ActivateRAToolsTab()
    ' 使用 Now 参数在当前队列操作完成后立即执行
    Application.OnTime Now, "DoActivateTab"

    ' 确保监听器存活
    If mAppEvents Is Nothing Then InitEvents
End Sub

' 延迟后执行激活
Public Sub DoActivateTab()
    On Error Resume Next
    If Not gRibbonUI Is Nothing Then
        gRibbonUI.ActivateTab "RATool"
    End If
End Sub

' AutoExec 宏
Public Sub AutoExec()
    InitEvents
End Sub