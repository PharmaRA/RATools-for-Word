VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuickToolbar
   Caption         =   "RATools"
   ClientHeight    =   3912
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
Private Const TOOLBAR_HEIGHT As Single = 224
Private Const FRAMELESS_CLIENT_WIDTH As Single = 127
Private Const HEADER_HEIGHT As Single = 22
Private Const EXPANDED_CLIENT_HEIGHT As Single = 192
Private Const BUTTON_TOP As Single = 26
Private Const BUTTON_WIDTH As Single = 26
Private Const BUTTON_HEIGHT As Single = 24
Private Const BUTTON_ROW_STEP As Single = 25
Private Const BUTTON_COLUMN_GAP As Single = 3
Private Const BUTTON_GROUP_GAP As Single = 4

Private m_HoverButtons As Collection
Private m_Dividers As Collection
Private m_ThemeApplied As Boolean
Private m_LastDark As Boolean
Private m_IsCollapsed As Boolean
Private m_IsFrameless As Boolean

Private WithEvents m_HdrBg As MSForms.Label
Private WithEvents m_HdrGrip As MSForms.Label
Private WithEvents m_HdrTitle As MSForms.Label
Private m_HdrDivider As MSForms.Label
Private m_BtnCollapse As clsFlatButton
Private m_BtnClose As clsFlatButton

Private Sub UserForm_Initialize()
    Me.Caption = "RATools"
    Me.Width = TOOLBAR_WIDTH
    Me.Height = TOOLBAR_HEIGHT
    m_IsCollapsed = False
    m_IsFrameless = False
    Set m_HoverButtons = New Collection

    CreateHeaderControls

    ConfigureButton Me.btnHeading1, 0, 0, 4, 0
    ConfigureButton Me.btnHeading2, 0, 1, 4, 1
    ConfigureButton Me.btnHeading3, 0, 2, 4, 2
    ConfigureButton Me.btnHeading4, 0, 3, 4, 3
    ConfigureButton Me.btnUnnumberedHeading1, 1, 0, 4, 4
    ConfigureButton Me.btnUnnumberedHeading2, 1, 1, 4, 5
    ConfigureButton Me.btnUnnumberedHeading3, 1, 2, 4, 6
    ConfigureButton Me.btnUnnumberedHeading4, 1, 3, 4, 7
    ConfigureButton Me.btnBodyText, 2, 0, 4, 8
    ConfigureButton Me.btnTableTitle, 2, 1, 4, 9
    ConfigureButton Me.btnFigureTitle, 2, 2, 4, 10
    ConfigureButton Me.btnPageBreakBefore, 2, 3, 4, 11
    ConfigureButton Me.btnFormatPainter, 3, 0, 4, 12
    ConfigureButton Me.btnParagraphSettings, 3, 1, 4, 13
    ConfigureButton Me.btnClearFormatting, 3, 2, 4, 14
    ConfigureButton Me.btnTextBlue, 3, 3, 4, 15
    ConfigureButton Me.btnAutoFitTable, 4, 0, 4, 16
    ConfigureButton Me.btnInsertCaption, 4, 1, 4, 17
    ConfigureButton Me.btnInsertCrossReference, 4, 2, 4, 18
    ConfigureButton Me.btnUpdateFields, 4, 3, 4, 19
    ConfigureButton Me.btnPasteRelink, 5, 0, 4, 20
    ConfigureButton Me.btnHyperlinksFieldsBlue, 5, 1, 4, 21
    ConfigureButton Me.btnAcceptRevisionsComments, 5, 2, 4, 22
    ConfigureButton Me.btnDetectHighlights, 5, 3, 4, 23

    CreateGroupDividers
    ApplyTheme True
    LayoutToolbarControls
    RestoreToolbarPosition
End Sub

Private Sub UserForm_Activate()
    ApplyTheme True
End Sub

Private Sub CreateHeaderControls()
    Set m_HdrBg = Me.Controls.Add("Forms.Label.1", "hdrBg", True)
    With m_HdrBg
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
    End With

    Set m_HdrGrip = Me.Controls.Add("Forms.Label.1", "hdrGrip", True)
    With m_HdrGrip
        .Caption = ChrW$(&HE784)
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .Font.Name = ThemeIconFontName()
        .Font.Size = 8
        .ControlTipText = "按住拖动悬浮工具栏"
    End With

    Set m_HdrTitle = Me.Controls.Add("Forms.Label.1", "hdrTitle", True)
    With m_HdrTitle
        .Caption = "RATools"
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
        .Font.Bold = True
        .ControlTipText = "按住拖动（双击可折叠/展开）"
    End With

    Set m_HdrDivider = Me.Controls.Add("Forms.Label.1", "hdrDivider", True)
    With m_HdrDivider
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
    End With

    Set m_BtnCollapse = CreateFlatButton(m_HoverButtons, Me, "hdrCollapse", vbNullString, "header", "collapse", Me, 8)
    m_BtnCollapse.SetIcon ChrW$(&HE70D), ThemeIconFontName(), 7.5
    m_BtnCollapse.SetToolTip "折叠工具栏"

    Set m_BtnClose = CreateFlatButton(m_HoverButtons, Me, "hdrClose", vbNullString, "close", "close", Me, 8)
    m_BtnClose.SetIcon ChrW$(&HE8BB), ThemeIconFontName(), 7.5
    m_BtnClose.SetToolTip "隐藏悬浮窗"
End Sub

Public Sub ApplyFramelessMode()
    MakeUserFormWindowFrameless Me.Caption
    m_IsFrameless = True
    ResizeForCurrentState
    LayoutToolbarControls
End Sub

Public Sub ToggleCollapsed()
    m_IsCollapsed = Not m_IsCollapsed
    If Not m_BtnCollapse Is Nothing Then
        If m_IsCollapsed Then
            m_BtnCollapse.SetIcon ChrW$(&HE70E), ThemeIconFontName(), 7.5
            m_BtnCollapse.SetToolTip "展开工具栏"
        Else
            m_BtnCollapse.SetIcon ChrW$(&HE70D), ThemeIconFontName(), 7.5
            m_BtnCollapse.SetToolTip "折叠工具栏"
        End If
    End If
    ResizeForCurrentState
    LayoutToolbarControls
    ReturnFocusToDocumentWindow
End Sub

Private Sub ResizeForCurrentState()
    Dim chromeW As Single
    Dim chromeH As Single
    Dim targetClientH As Single

    If Not m_IsFrameless Then Exit Sub

    On Error Resume Next
    chromeW = Me.Width - Me.InsideWidth
    chromeH = Me.Height - Me.InsideHeight
    If chromeW < 0 Then chromeW = 0
    If chromeH < 0 Then chromeH = 0

    If m_IsCollapsed Then
        targetClientH = HEADER_HEIGHT
    Else
        targetClientH = EXPANDED_CLIENT_HEIGHT
    End If

    Me.Width = FRAMELESS_CLIENT_WIDTH + chromeW
    Me.Height = targetClientH + chromeH
    On Error GoTo 0
End Sub

Private Sub LayoutToolbarControls()
    Dim cw As Single
    Dim ctrl As Object
    Dim divCtrl As MSForms.Label

    On Error Resume Next
    cw = Me.InsideWidth
    If cw <= 0 Then cw = FRAMELESS_CLIENT_WIDTH

    If Not m_HdrBg Is Nothing Then m_HdrBg.Move 0, 0, cw, HEADER_HEIGHT - 0.75
    If Not m_HdrGrip Is Nothing Then m_HdrGrip.Move 4, (HEADER_HEIGHT - 11) / 2, 10, 11
    If Not m_HdrTitle Is Nothing Then m_HdrTitle.Move 15, (HEADER_HEIGHT - 12.5) / 2, cw - 54, 12.5
    If Not m_BtnCollapse Is Nothing Then m_BtnCollapse.MoveTo cw - 38, 2, 17, HEADER_HEIGHT - 4
    If Not m_BtnClose Is Nothing Then m_BtnClose.MoveTo cw - 19, 2, 17, HEADER_HEIGHT - 4
    If Not m_HdrDivider Is Nothing Then
        m_HdrDivider.Move 0, HEADER_HEIGHT - 0.75, cw, 0.75
        m_HdrDivider.Visible = (Not m_IsCollapsed)
    End If

    If Not m_HdrBg Is Nothing Then m_HdrBg.ZOrder 1
    If Not m_HdrGrip Is Nothing Then m_HdrGrip.ZOrder 0
    If Not m_HdrTitle Is Nothing Then m_HdrTitle.ZOrder 0
    If Not m_BtnCollapse Is Nothing Then m_BtnCollapse.MoveTo cw - 38, 2, 17, HEADER_HEIGHT - 4
    If Not m_BtnClose Is Nothing Then m_BtnClose.MoveTo cw - 19, 2, 17, HEADER_HEIGHT - 4
    If Not m_HdrDivider Is Nothing Then m_HdrDivider.ZOrder 0

    For Each ctrl In Me.Controls
        If Left$(ctrl.Name, 3) = "btn" Then
            ctrl.Visible = (Not m_IsCollapsed)
        End If
    Next ctrl

    If Not m_Dividers Is Nothing Then
        For Each divCtrl In m_Dividers
            divCtrl.Visible = (Not m_IsCollapsed)
        Next divCtrl
    End If
    On Error GoTo 0
End Sub

Public Sub ApplyTheme(Optional ByVal force As Boolean = False)
    Dim darkNow As Boolean
    Dim divCtrl As MSForms.Label

    darkNow = IsOfficeDarkTheme()
    If m_ThemeApplied And Not force And (darkNow = m_LastDark) Then Exit Sub

    ApplyFormBaseTheme Me

    On Error Resume Next
    If Not m_HdrBg Is Nothing Then m_HdrBg.BackColor = ThemeCanvasColor()
    If Not m_HdrGrip Is Nothing Then m_HdrGrip.ForeColor = ThemeTextMutedColor()
    If Not m_HdrTitle Is Nothing Then m_HdrTitle.ForeColor = ThemeTextPrimaryColor()
    If Not m_HdrDivider Is Nothing Then StyleDividerLine m_HdrDivider
    On Error GoTo 0

    If Not m_Dividers Is Nothing Then
        For Each divCtrl In m_Dividers
            StyleDividerLine divCtrl
        Next divCtrl
    End If
    RefreshHoverButtons m_HoverButtons
    ApplyFormWindowChrome Me.Caption, False

    m_LastDark = darkNow
    m_ThemeApplied = True
End Sub

Public Sub OnHoverControlChanged(ByVal activeItem As Object)
    ResetHoverButtons m_HoverButtons, activeItem
    RefreshTooltipWatch
End Sub

Public Sub OnFlatButtonClick(ByVal actionKey As String)
    Select Case actionKey
        Case "collapse"
            ToggleCollapsed
        Case "close"
            SaveToolbarPosition
            Me.Hide
    End Select
End Sub

Private Sub m_HdrBg_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    DragUserFormWindow Me.Caption
    SaveToolbarPosition
End Sub

Private Sub m_HdrGrip_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    DragUserFormWindow Me.Caption
    SaveToolbarPosition
End Sub

Private Sub m_HdrTitle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    DragUserFormWindow Me.Caption
    SaveToolbarPosition
End Sub

Private Sub m_HdrBg_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ToggleCollapsed
End Sub

Private Sub m_HdrTitle_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ToggleCollapsed
End Sub

Private Sub m_HdrBg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    RefreshTooltipWatch
End Sub

Private Sub m_HdrGrip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    RefreshTooltipWatch
End Sub

Private Sub m_HdrTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    RefreshTooltipWatch
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
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
        .BackColor = ThemeToolbarTileColor()
        .BorderStyle = fmBorderStyleSingle
        .borderColor = ThemeToolbarTileBorderColor()
        .SpecialEffect = fmSpecialEffectFlat
        .PicturePosition = fmPicturePositionCenter
        .TextAlign = fmTextAlignCenter
        .tabIndex = tabIndex
        .Font.Name = "Microsoft YaHei"
        .Font.Size = 7
    End With

    If HasButtonPicture(buttonItem) Then buttonItem.Caption = vbNullString
    RegisterHoverButton m_HoverButtons, buttonItem, "toolbar", Me
End Sub

Private Sub CreateGroupDividers()
    Dim gapRows As Variant
    Dim idx As Long
    Dim gapRow As Long
    Dim prevBottom As Single
    Dim nextTop As Single
    Dim rowWidth As Single
    Dim divCtrl As MSForms.Label

    Set m_Dividers = New Collection
    gapRows = Array(2, 4, 5)
    rowWidth = 4 * BUTTON_WIDTH + 3 * BUTTON_COLUMN_GAP

    On Error Resume Next
    For idx = LBound(gapRows) To UBound(gapRows)
        gapRow = CLng(gapRows(idx))
        prevBottom = GetButtonTop(gapRow - 1) + BUTTON_HEIGHT
        nextTop = GetButtonTop(gapRow)
        Set divCtrl = Me.Controls.Add("Forms.Label.1", "divGroup" & CStr(idx + 1), True)
        If Not divCtrl Is Nothing Then
            divCtrl.Move GetButtonLeft(0, 4), prevBottom + (nextTop - prevBottom - 0.75) / 2, rowWidth, 0.75
            StyleDividerLine divCtrl
            m_Dividers.Add divCtrl
        End If
    Next idx
    On Error GoTo 0
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
    If rowIndex >= 5 Then GetButtonTop = GetButtonTop + BUTTON_GROUP_GAP
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
Private Sub btnBodyText_Click()
    RunQuickToolbarAction "style_body_text"
End Sub

Private Sub btnBodyText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
Private Sub btnPageBreakBefore_Click()
    RunQuickToolbarAction "page_break_before"
End Sub

Private Sub btnPageBreakBefore_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
Private Sub btnClearFormatting_Click()
    RunQuickToolbarAction "clear_formatting"
End Sub

Private Sub btnClearFormatting_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnTextBlue_Click()
    RunQuickToolbarAction "text_blue"
End Sub

Private Sub btnTextBlue_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnAutoFitTable_Click()
    RunQuickToolbarAction "autofit_table"
End Sub

Private Sub btnAutoFitTable_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
Private Sub btnInsertCaption_Click()
    RunQuickToolbarAction "insert_caption"
End Sub

Private Sub btnInsertCaption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
Private Sub btnPasteRelink_Click()
    RunQuickToolbarAction "paste_relink_cross_reference"
End Sub

Private Sub btnPasteRelink_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
    ClearHoverButtons m_HoverButtons
    Set m_Dividers = Nothing
    Set m_BtnCollapse = Nothing
    Set m_BtnClose = Nothing
    Set m_HdrDivider = Nothing
    Set m_HdrTitle = Nothing
    Set m_HdrGrip = Nothing
    Set m_HdrBg = Nothing
End Sub

Private Sub UserForm_Terminate()
    SaveToolbarPosition
End Sub
