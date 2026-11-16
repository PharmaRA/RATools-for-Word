Attribute VB_Name = "UI_ReviewToolbar"
Option Explicit

Public Const REVIEW_TOOLBAR_CAPTION As String = "RATools 审阅工具"
Private mActionRunning As Boolean

' 只查询已加载实例；应用事件不得通过默认实例隐式加载窗体。
Private Function LoadedReviewToolbar() As Object
    Dim formItem As Object
    For Each formItem In VBA.UserForms
        If TypeName(formItem) = "frmReviewToolbar" Then
            Set LoadedReviewToolbar = formItem
            Exit Function
        End If
    Next formItem
End Function

Public Sub ToggleReviewToolbar(ByVal control As IRibbonControl)
    Dim toolbar As Object
    On Error GoTo Failed
    Set toolbar = LoadedReviewToolbar()
    If Not toolbar Is Nothing Then
        If toolbar.Visible Then
            toolbar.SaveToolbarPosition
            toolbar.Hide
            Exit Sub
        End If
    End If
    ShowReviewToolbar
    Exit Sub
Failed:
    MsgBox "无法切换审阅悬浮窗：" & Err.Description, vbExclamation, REVIEW_TOOLBAR_CAPTION
End Sub

Public Sub ShowReviewToolbar()
    On Error GoTo Failed
    UI_Ribbon.InitEvents
    Load frmReviewToolbar
    frmReviewToolbar.Show vbModeless
    DetachUserFormWindowFromOwner REVIEW_TOOLBAR_CAPTION
    RefreshReviewToolbarIfVisible
    ReturnFocusToDocumentWindow
    Exit Sub
Failed:
    MsgBox "无法显示审阅悬浮窗：" & Err.Description, vbExclamation, REVIEW_TOOLBAR_CAPTION
End Sub

Public Sub RefreshReviewToolbarIfVisible()
    Dim toolbar As Object
    On Error Resume Next
    Set toolbar = LoadedReviewToolbar()
    If toolbar Is Nothing Then Exit Sub
    If Not toolbar.Visible Then Exit Sub
    toolbar.RefreshDocumentState mActionRunning
    KeepUserFormWindowOnTop REVIEW_TOOLBAR_CAPTION
End Sub

Public Sub UnloadReviewToolbarOnQuit()
    Dim toolbar As Object
    On Error Resume Next
    Set toolbar = LoadedReviewToolbar()
    If toolbar Is Nothing Then Exit Sub
    toolbar.SaveToolbarPosition
    Unload toolbar
End Sub

Public Sub RefreshReviewToolbarTooltip()
    On Error Resume Next
    If Not IsUserFormWindowVisible(REVIEW_TOOLBAR_CAPTION) Then Exit Sub
    EnsureTooltipOnTop REVIEW_TOOLBAR_CAPTION
End Sub

Public Sub RunReviewToolbarAction(ByVal actionKey As String)
    Dim targetDocument As Document
    Dim keepInteractionFocus As Boolean
    Dim failureText As String

    If mActionRunning Then Exit Sub
    On Error GoTo Failed
    If Documents.Count = 0 Then
        RefreshReviewToolbarIfVisible
        Exit Sub
    End If
    ' 每次执行重新取当前文档，不保留可能因格式处理重开而失效的引用。
    Set targetDocument = ActiveDocument
    actionKey = LCase$(Trim$(actionKey))
    mActionRunning = True
    RefreshReviewToolbarIfVisible

    Select Case actionKey
        Case "track_changes"
            targetDocument.TrackRevisions = Not targetDocument.TrackRevisions
        Case "new_comment"
            ExecuteQuickToolbarMso "ReviewNewComment"
            keepInteractionFocus = True
        Case "reviewer_filter"
            Wrapper_ShowReviewerFilter
            keepInteractionFocus = True
        Case "accept_format_revisions"
            Wrapper_AcceptFormatRevisions
        Case "detect_highlights"
            BatchDetectHighlights
        Case "detect_clean_versions"
            BatchDetectCleanVersions
        Case "accept_revisions_comments"
            BatchAcceptAndClean
        Case Else
            Err.Raise vbObjectError + 2050, "RunReviewToolbarAction", "未知的审阅操作：" & actionKey
    End Select

Finished:
    mActionRunning = False
    Set targetDocument = Nothing
    RefreshReviewToolbarIfVisible
    If Not keepInteractionFocus Then ReturnFocusToDocumentWindow
    If Len(failureText) > 0 Then MsgBox failureText, vbExclamation, REVIEW_TOOLBAR_CAPTION
    Exit Sub
Failed:
    failureText = "审阅操作执行失败：" & Err.Description
    Resume Finished
End Sub
