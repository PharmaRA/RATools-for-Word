Attribute VB_Name = "Mod_Core_Session"
Option Explicit

' =============================================
' 公共层：批处理会话守卫与文档安全开关
' 约定：入口 BeginBatchUI，所有出口（含错误处理路径）EndBatchUI 成对调用
' =============================================

Public Sub BeginBatchUI(Optional ByVal suppressAlerts As Boolean = False)
    Application.ScreenUpdating = False
    If suppressAlerts Then Application.DisplayAlerts = wdAlertsNone
End Sub

' 恢复屏幕刷新/警告/状态栏；容错设计，可在任何状态下安全调用
Public Sub EndBatchUI()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' 打开文档；若已在 Word 中打开则复用现有实例（wasAlreadyOpen 返回 True）。
' 打开失败返回 Nothing，由调用方决定如何计数或提示。
Public Function SafeOpenDocument(ByVal filePath As String, _
                                 ByVal openReadOnly As Boolean, _
                                 ByRef wasAlreadyOpen As Boolean, _
                                 Optional ByRef errorMessage As String) As Document
    Dim doc As Document

    wasAlreadyOpen = False
    errorMessage = ""

    On Error Resume Next
    Set doc = Documents(filePath)
    On Error GoTo 0

    If Not doc Is Nothing Then
        wasAlreadyOpen = True
        Set SafeOpenDocument = doc
        Exit Function
    End If

    On Error Resume Next
    Set doc = Documents.Open(fileName:=filePath, ReadOnly:=openReadOnly, _
                             Visible:=False, AddToRecentFiles:=False)
    If Err.Number <> 0 Then errorMessage = Err.Description
    On Error GoTo 0

    Set SafeOpenDocument = doc
End Function

' 静默关闭文档（可选先保存），出错不抛出
Public Sub CloseDocumentQuietly(ByVal doc As Document, ByVal saveChanges As Boolean)
    If doc Is Nothing Then Exit Sub
    On Error Resume Next
    If saveChanges Then doc.Save
    doc.Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
End Sub