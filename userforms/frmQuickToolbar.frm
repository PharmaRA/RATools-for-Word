VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuickToolbar
   Caption         =   "RATools"
   ClientHeight    =   2796
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2460
   OleObjectBlob   =   "frmQuickToolbar.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmQuickToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SETTINGS_APP As String = "RATools"
Private Const SETTINGS_SECTION As String = "QuickToolbar"
Private Const TOOLBAR_WIDTH As Single = 134
Private Const TOOLBAR_HEIGHT As Single = 168
Private Const BUTTON_TOP As Single = 3
Private Const BUTTON_WIDTH As Single = 26
Private Const BUTTON_HEIGHT As Single = 24
Private Const BUTTON_ROW_STEP As Single = 25
Private Const BUTTON_COLUMN_GAP As Single = 3
Private Const BUTTON_GROUP_GAP As Single = 4

Private Sub UserForm_Initialize()
    Me.Caption = "RATools"
    Me.Width = TOOLBAR_WIDTH
    Me.Height = TOOLBAR_HEIGHT
    Me.BackColor = RGB(247, 248, 250)

    ConfigureButton Me.btnHeading1, 0, 0, 4, 0
    ConfigureButton Me.btnHeading2, 0, 1, 4, 1
    ConfigureButton Me.btnHeading3, 0, 2, 4, 2
    ConfigureButton Me.btnHeading4, 0, 3, 4, 3
    ConfigureButton Me.btnUnnumberedHeading1, 1, 0, 4, 4
    ConfigureButton Me.btnUnnumberedHeading2, 1, 1, 4, 5
    ConfigureButton Me.btnUnnumberedHeading3, 1, 2, 4, 6
    ConfigureButton Me.btnUnnumberedHeading4, 1, 3, 4, 7
    ConfigureButton Me.btnFormatPainter, 2, 0, 4, 8
    ConfigureButton Me.btnParagraphSettings, 2, 1, 4, 9
    ConfigureButton Me.btnPageBreakBefore, 2, 2, 4, 10
    ConfigureButton Me.btnTableTitle, 2, 3, 4, 11
    ConfigureButton Me.btnFigureTitle, 3, 0, 4, 12
    ConfigureButton Me.btnAutoFitTable, 3, 1, 4, 13
    ConfigureButton Me.btnInsertCrossReference, 3, 2, 4, 14
    ConfigureButton Me.btnUpdateFields, 3, 3, 4, 15
    ConfigureButton Me.btnHyperlinksFieldsBlue, 4, 0, 3, 16
    ConfigureButton Me.btnAcceptRevisionsComments, 4, 1, 3, 17
    ConfigureButton Me.btnDetectHighlights, 4, 2, 3, 18

    RestoreToolbarPosition
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub

Private Sub ConfigureButton(ByVal buttonItem As MSForms.Label, _
                            ByVal rowIndex As Long, _
                            ByVal columnIndex As Long, _
                            ByVal columnCount As Long, _
                            ByVal tabIndex As Long)
    With buttonItem
        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT
        .Left = GetButtonLeft(columnIndex, columnCount)
        .Top = GetButtonTop(rowIndex)
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(221, 226, 232)
        .PicturePosition = fmPicturePositionCenter
        .TextAlign = fmTextAlignCenter
        .tabIndex = tabIndex
        .Font.Name = "Microsoft YaHei"
        .Font.Size = 7
    End With

    If HasButtonPicture(buttonItem) Then buttonItem.Caption = vbNullString
End Sub

Private Function GetButtonLeft(ByVal columnIndex As Long, _
                               ByVal columnCount As Long) As Single
    Dim rowWidth As Single

    rowWidth = columnCount * BUTTON_WIDTH + _
               (columnCount - 1) * BUTTON_COLUMN_GAP
    GetButtonLeft = (Me.InsideWidth - rowWidth) / 2 + _
                    columnIndex * (BUTTON_WIDTH + BUTTON_COLUMN_GAP)
End Function

Private Function GetButtonTop(ByVal rowIndex As Long) As Single
    GetButtonTop = BUTTON_TOP + rowIndex * BUTTON_ROW_STEP
    If rowIndex >= 2 Then GetButtonTop = GetButtonTop + BUTTON_GROUP_GAP
    If rowIndex >= 4 Then GetButtonTop = GetButtonTop + BUTTON_GROUP_GAP
End Function

Private Function HasButtonPicture(ByVal buttonItem As MSForms.Label) As Boolean
    Dim pictureItem As Object

    On Error GoTo ErrH
    Set pictureItem = buttonItem.Picture
    HasButtonPicture = Not (pictureItem Is Nothing)
    Exit Function

ErrH:
    HasButtonPicture = False
End Function

Private Sub btnHeading1_Click()
    RunQuickToolbarAction "style_heading1"
End Sub

Private Sub btnHeading1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnHeading2_Click()
    RunQuickToolbarAction "style_heading2"
End Sub

Private Sub btnHeading2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnHeading3_Click()
    RunQuickToolbarAction "style_heading3"
End Sub

Private Sub btnHeading3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnHeading4_Click()
    RunQuickToolbarAction "style_heading4"
End Sub

Private Sub btnHeading4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnUnnumberedHeading1_Click()
    RunQuickToolbarAction "style_unnumbered_heading1"
End Sub

Private Sub btnUnnumberedHeading1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnUnnumberedHeading2_Click()
    RunQuickToolbarAction "style_unnumbered_heading2"
End Sub

Private Sub btnUnnumberedHeading2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnUnnumberedHeading3_Click()
    RunQuickToolbarAction "style_unnumbered_heading3"
End Sub

Private Sub btnUnnumberedHeading3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnUnnumberedHeading4_Click()
    RunQuickToolbarAction "style_unnumbered_heading4"
End Sub

Private Sub btnUnnumberedHeading4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnFormatPainter_Click()
    RunQuickToolbarAction "format_painter"
End Sub

Private Sub btnFormatPainter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnParagraphSettings_Click()
    RunQuickToolbarAction "paragraph_settings"
End Sub

Private Sub btnParagraphSettings_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnPageBreakBefore_Click()
    RunQuickToolbarAction "page_break_before"
End Sub

Private Sub btnPageBreakBefore_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnTableTitle_Click()
    RunQuickToolbarAction "style_table_title"
End Sub

Private Sub btnTableTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnFigureTitle_Click()
    RunQuickToolbarAction "style_figure_title"
End Sub

Private Sub btnFigureTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnAutoFitTable_Click()
    RunQuickToolbarAction "autofit_table"
End Sub

Private Sub btnAutoFitTable_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnInsertCrossReference_Click()
    RunQuickToolbarAction "insert_cross_reference"
End Sub

Private Sub btnInsertCrossReference_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnUpdateFields_Click()
    RunQuickToolbarAction "update_fields"
End Sub

Private Sub btnUpdateFields_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnHyperlinksFieldsBlue_Click()
    RunQuickToolbarAction "hyperlinks_fields_blue"
End Sub

Private Sub btnHyperlinksFieldsBlue_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnAcceptRevisionsComments_Click()
    RunQuickToolbarAction "accept_revisions_comments"
End Sub

Private Sub btnAcceptRevisionsComments_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnDetectHighlights_Click()
    RunQuickToolbarAction "detect_highlights"
End Sub

Private Sub btnDetectHighlights_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub

Private Sub RestoreToolbarPosition()
    Dim savedLeft As String
    Dim savedTop As String

    savedLeft = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Left", "")
    savedTop = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Top", "")

    If IsNumeric(savedLeft) And IsNumeric(savedTop) Then
        Me.Left = CSng(savedLeft)
        Me.Top = CSng(savedTop)
    Else
        PositionNearWordWindow
    End If
End Sub

Private Sub PositionNearWordWindow()
    On Error GoTo UseFallback

    Me.Left = Application.Left + Application.Width - Me.Width - 24
    Me.Top = Application.Top + 55
    Exit Sub

UseFallback:
    Me.Left = 80
    Me.Top = 55
End Sub

Public Sub SaveToolbarPosition()
    On Error Resume Next
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Left", CStr(Me.Left)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Top", CStr(Me.Top)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveToolbarPosition
End Sub

Private Sub UserForm_Terminate()
    SaveToolbarPosition
End Sub
