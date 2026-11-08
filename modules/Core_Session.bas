Attribute VB_Name = "Core_Session"
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

Public Function TrySaveDocument(ByVal doc As Document, _
                                Optional ByRef errorMessage As String) As Boolean
    errorMessage = ""
    If doc Is Nothing Then
        errorMessage = "文档不存在。"
        Exit Function
    End If
    On Error GoTo Failed
    doc.Save
    If Not doc.Saved Then
        errorMessage = "保存被取消或未完成。"
        GoTo KeepOpen
    End If
    TrySaveDocument = True
    Exit Function
Failed:
    errorMessage = Err.Description
KeepOpen:
    On Error Resume Next
    ' Interactive Word must expose the retained document; hidden automation
    ' keeps control of its own application visibility.
    If Application.Visible Then
        If doc.Windows.Count > 0 Then doc.Windows(1).Visible = True
    End If
    On Error GoTo 0
End Function

' 保存失败时保留文档；返回值供批处理统计使用。
Public Function CloseDocumentQuietly(ByVal doc As Document, ByVal saveChanges As Boolean, _
                                     Optional ByRef errorMessage As String) As Boolean
    Dim openDoc As Document
    errorMessage = ""
    If doc Is Nothing Then Exit Function
    If saveChanges Then
        If Not TrySaveDocument(doc, errorMessage) Then Exit Function
    End If
    On Error GoTo Failed
    doc.Close SaveChanges:=wdDoNotSaveChanges
    For Each openDoc In Application.Documents
        If openDoc Is doc Then
            errorMessage = "关闭被取消，文档保持打开。"
            GoTo KeepOpen
        End If
    Next openDoc
    CloseDocumentQuietly = True
    Exit Function
Failed:
    errorMessage = Err.Description
KeepOpen:
    On Error Resume Next
    ' Interactive Word must expose the retained document; hidden automation
    ' keeps control of its own application visibility.
    If Application.Visible Then
        If doc.Windows.Count > 0 Then doc.Windows(1).Visible = True
    End If
End Function
