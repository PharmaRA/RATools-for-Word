Attribute VB_Name = "Mod_QuickToolbarActions"
Option Explicit

' =============================================
' 悬浮工具栏 MSO 命令执行
' 分发表已于 v0.8 阶段6合并进 Mod_QuickToolbar.RunQuickToolbarAction，
' 本模块只保留 ExecuteMso 包装与其测试模式。
' =============================================

' ====== TEST SUPPORT：测试模式状态（仅测试脚本切换） ======
Private mQuickToolbarMsoTestMode As Boolean
Private mLastQuickToolbarMso As String

Public Sub SetQuickToolbarMsoTestMode(ByVal enabled As Boolean)
    mQuickToolbarMsoTestMode = enabled
    mLastQuickToolbarMso = vbNullString
End Sub

Public Function GetLastQuickToolbarMsoForTest() As String
    GetLastQuickToolbarMsoForTest = mLastQuickToolbarMso
End Function

' 执行 Word 内置命令；测试模式下仅记录命令 ID 不实际执行
Public Sub ExecuteQuickToolbarMso(ByVal commandId As String)
    If mQuickToolbarMsoTestMode Then
        mLastQuickToolbarMso = commandId
    Else
        Application.CommandBars.ExecuteMso commandId
    End If
End Sub