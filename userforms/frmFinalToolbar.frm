VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinalToolbar 
   Caption         =   "RATools 定稿工具"
   ClientHeight    =   7680
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8424.001
   OleObjectBlob   =   "frmFinalToolbar.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmFinalToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mUpdating As Boolean
Private mMenuRange As Range
Private mSnapshotDocument As Document
Private mSnapshotState As String, mSnapshotFailure As String
Private mSnapshotStale As Boolean, mSnapshotSaved As Boolean, mTracking As Boolean
Private mRevisionCount As Long, mCommentCount As Long, mReadAt As Date
Private Sub UserForm_Initialize()
    Dim item As Object
    For Each item In Me.Controls: item.Font.Name = "Microsoft YaHei": item.Font.Size = 9: Next item
    On Error Resume Next
    Me.Left = CSng(GetSetting("RATools", "FinalToolbar", "Left", CStr(Application.Left + 30)))
    Me.Top = CSng(GetSetting("RATools", "FinalToolbar", "Top", CStr(Application.Top + 80)))
    On Error GoTo 0
    InitializeMenus
    RefreshDocumentState False
End Sub
Private Sub btn0_Click()
    RunWorkflowAction "detect_clean_versions"
End Sub
Private Sub btn1_Click()
    RunWorkflowAction "detect_highlights"
End Sub
Private Sub btn2_Click()
    RunWorkflowAction "mso:TableOfContentsUpdate"
End Sub
Private Sub btn3_Click()
    RunWorkflowAction "mso:TableOfFiguresUpdate"
End Sub
Private Sub btn4_Click()
    RunWorkflowAction "mso:FieldsUpdate"
End Sub
Private Sub btn5_Click()
    RunWorkflowAction "protect_field_format"
End Sub
Private Sub btn6_Click()
    RunWorkflowAction "accept_revisions_comments"
End Sub
Private Sub btn7_Click()
    RunWorkflowAction "heading_numbers_text"
End Sub
Private Sub btn8_Click()
    RunWorkflowAction "save_current"
End Sub
Private Sub btn9_Click()
    RunWorkflowAction "export_pdf"
End Sub
Private Sub btn10_Click()
    RunWorkflowAction "rename_current"
End Sub
Private Sub btn11_Click()
    RunWorkflowAction "open_folder"
End Sub
Private Sub InitializeMenus()
    mUpdating = True
    mUpdating = False
End Sub
Private Sub ResetSnapshot()
    Set mSnapshotDocument = Nothing
    mSnapshotState = "unread"
    mSnapshotStale = False
    mSnapshotFailure = ""
End Sub
Private Sub RenderSnapshot()
    Select Case mSnapshotState
        Case "read"
            lblSnapshot.Caption = IIf(mSnapshotStale, "上次读数（待刷新）：", "") & "修订：" & mRevisionCount & "    批注：" & mCommentCount & "    修订跟踪：" & IIf(mTracking, "开", "关")
            lblSnapshotTime.Caption = "读取时间：" & Format$(mReadAt, "hh:nn:ss") & IIf(mSnapshotStale, " · 待刷新", "")
        Case "failed"
            lblSnapshot.Caption = "状态读取失败：" & mSnapshotFailure
            lblSnapshotTime.Caption = "状态读取时间：—"
        Case Else
            lblSnapshot.Caption = "修订：—    批注：—    修订跟踪：—"
            lblSnapshotTime.Caption = "状态读取时间：—"
    End Select
End Sub
Public Sub ReadDocumentSnapshot()
    Dim values As Variant
    On Error GoTo Failed
    If Documents.Count = 0 Then Call ResetSnapshot: RenderSnapshot: Exit Sub
    Set mSnapshotDocument = ActiveDocument
    values = ReadWorkflowReviewSnapshot(mSnapshotDocument)
    mRevisionCount = CLng(values(0)): mCommentCount = CLng(values(1)): mTracking = CBool(values(2))
    mSnapshotSaved = mSnapshotDocument.Saved
    mReadAt = Now
    mSnapshotState = "read": mSnapshotStale = False: mSnapshotFailure = ""
    RenderSnapshot
    Exit Sub
Failed:
    mSnapshotState = "failed": mSnapshotStale = False
    mSnapshotFailure = Err.Description
    RenderSnapshot
End Sub
Public Sub RefreshDocumentState(ByVal busy As Boolean, Optional ByVal mayHaveEdited As Boolean = False)
    Dim item As Object, available As Boolean, needsRead As Boolean
    On Error GoTo Unavailable
    If Documents.Count = 0 Then
        lblDocument.Caption = "当前没有打开的文档"
        Call ResetSnapshot: RenderSnapshot
        GoTo UpdateControls
    End If
    lblDocument.Caption = "当前文档：" & ActiveDocument.Name & IIf(ActiveDocument.Saved, " · 已保存", " · 有未保存修改")
    available = True
    If mSnapshotDocument Is Nothing Then
        needsRead = True
    ElseIf Not (mSnapshotDocument Is ActiveDocument) Then
        ResetSnapshot
        needsRead = True
    End If
    If needsRead And Not busy Then ReadDocumentSnapshot
    If mSnapshotState = "read" Then
        If mayHaveEdited Or (mSnapshotSaved <> ActiveDocument.Saved) Then mSnapshotStale = True
    End If
    RenderSnapshot
    GoTo UpdateControls
Unavailable:
    available = False
    lblDocument.Caption = "当前文档不可用"
    Call ResetSnapshot: RenderSnapshot
UpdateControls:
    For Each item In Me.Controls
        If TypeName(item) = "CommandButton" Or TypeName(item) = "ComboBox" Then item.Enabled = available And Not busy
    Next item
End Sub
Private Sub btnRefreshState_Click()
    RunWorkflowAction "refresh_final_state"
End Sub
Private Sub UserForm_Activate(): RefreshWorkflowToolbarsIfVisible: End Sub
Public Sub SaveToolbarPosition()
    On Error Resume Next
    SaveSetting "RATools", "FinalToolbar", "Left", CStr(Me.Left)
    SaveSetting "RATools", "FinalToolbar", "Top", CStr(Me.Top)
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer): SaveToolbarPosition: End Sub
Private Sub lblDocument_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblSnapshot_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblSnapshotTime_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btnRefreshState_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblGroup0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblGroup2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblGroup6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblGroup8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
