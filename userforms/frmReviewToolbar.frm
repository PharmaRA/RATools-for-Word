VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReviewToolbar 
   Caption         =   "RATools 审阅工具"
   ClientHeight    =   7932
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4548
   OleObjectBlob   =   "frmReviewToolbar.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmReviewToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SETTINGS_SECTION As String = "ReviewToolbar"

Private Sub UserForm_Initialize()
    Dim controlItem As Object
    For Each controlItem In Me.Controls
        controlItem.Font.Name = "Microsoft YaHei"
        controlItem.Font.Size = 9
    Next controlItem
    RestoreToolbarPosition
    RefreshDocumentState False
End Sub

Private Sub UserForm_Activate()
    RefreshReviewToolbarIfVisible
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Public Sub RefreshDocumentState(ByVal busy As Boolean)
    Dim controlItem As Object
    Dim hasDocument As Boolean
    Dim targetDocument As Document

    On Error GoTo Unavailable
    If Application.Documents.Count > 0 Then
        Set targetDocument = ActiveDocument
        Me.lblDocument.Caption = "当前文档：" & targetDocument.Name
        Me.lblDocument.ControlTipText = Me.lblDocument.Caption
        If targetDocument.TrackRevisions Then
            Me.btnTrackChanges.Caption = "修订跟踪：已开启"
        Else
            Me.btnTrackChanges.Caption = "修订跟踪：已关闭"
        End If
        hasDocument = True
    Else
        Me.lblDocument.Caption = "当前没有打开的文档"
        Me.lblDocument.ControlTipText = Me.lblDocument.Caption
        Me.btnTrackChanges.Caption = "修订跟踪：不可用"
    End If
    GoTo UpdateButtons
Unavailable:
    hasDocument = False
    Me.lblDocument.Caption = "当前文档不可用"
    Me.btnTrackChanges.Caption = "修订跟踪：不可用"
UpdateButtons:
    For Each controlItem In Me.Controls
        If TypeName(controlItem) = "CommandButton" Then controlItem.Enabled = hasDocument And Not busy
    Next controlItem
End Sub

Private Sub btnTrackChanges_Click()
    RunReviewToolbarAction "track_changes"
End Sub

Private Sub btnNewComment_Click()
    RunReviewToolbarAction "new_comment"
End Sub

Private Sub btnReviewerFilter_Click()
    RunReviewToolbarAction "reviewer_filter"
End Sub

Private Sub btnAcceptFormat_Click()
    RunReviewToolbarAction "accept_format_revisions"
End Sub

Private Sub btnDetectHighlights_Click()
    RunReviewToolbarAction "detect_highlights"
End Sub

Private Sub btnDetectClean_Click()
    RunReviewToolbarAction "detect_clean_versions"
End Sub

Private Sub btnClean_Click()
    RunReviewToolbarAction "accept_revisions_comments"
End Sub

Private Sub RestoreToolbarPosition()
    Dim savedLeft As String
    Dim savedTop As String
    On Error GoTo UseDefault
    savedLeft = GetSetting("RATools", SETTINGS_SECTION, "Left", "")
    savedTop = GetSetting("RATools", SETTINGS_SECTION, "Top", "")
    If IsNumeric(savedLeft) And IsNumeric(savedTop) Then
        Me.Left = CSng(savedLeft)
        Me.Top = CSng(savedTop)
        Exit Sub
    End If
UseDefault:
    Me.Left = Application.Left + Application.Width - Me.Width - 174
    Me.Top = Application.Top + 55
End Sub

Public Sub SaveToolbarPosition()
    On Error Resume Next
    SaveSetting "RATools", SETTINGS_SECTION, "Left", CStr(Me.Left)
    SaveSetting "RATools", SETTINGS_SECTION, "Top", CStr(Me.Top)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveToolbarPosition
End Sub


Private Sub lblDocument_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub lblDaily_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub btnTrackChanges_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub btnNewComment_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub btnReviewerFilter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub lblFormat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub btnAcceptFormat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub lblCheck_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub btnDetectHighlights_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub btnDetectClean_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub lblClean_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub btnClean_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

Private Sub lblWarning_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshReviewToolbarTooltip
End Sub

