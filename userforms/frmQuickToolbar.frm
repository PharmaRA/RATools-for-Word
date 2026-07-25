VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuickToolbar
   Caption         =   "RATools 快捷工具栏"
   ClientHeight    =   7000
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2800
   OleObjectBlob   =   "frmQuickToolbar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   0   'Manual
End
Attribute VB_Name = "frmQuickToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SETTINGS_APP As String = "RATools"
Private Const SETTINGS_SECTION As String = "QuickToolbar"
Private Const TOOLBAR_WIDTH As Single = 135
Private Const TOOLBAR_HEIGHT As Single = 370
Private Const BUTTON_LEFT As Single = 43
Private Const BUTTON_WIDTH As Single = 48
Private Const BUTTON_HEIGHT As Single = 34
Private Const BUTTON_STEP As Single = 40

Private Sub UserForm_Initialize()
    Me.Caption = "RATools 快捷工具栏"
    Me.Width = TOOLBAR_WIDTH
    Me.Height = TOOLBAR_HEIGHT
    Me.BackColor = RGB(245, 245, 245)

    ConfigureButton Me.btnHeading1, "标题 1", "应用“标题1-F”样式", 0
    ConfigureButton Me.btnHeading2, "标题 2", "应用“标题2-F”样式", 1
    ConfigureButton Me.btnHeading3, "标题 3", "应用“标题3-F”样式", 2
    ConfigureButton Me.btnBody, "正文", "应用“正文-F”样式", 3
    ConfigureButton Me.btnTextBlue, "设为蓝色", "将选中文字设置为标准蓝色", 4
    ConfigureButton Me.btnPageBreakBefore, "段前分页", "切换选中段落的段前分页属性", 5
    ConfigureButton Me.btnAutoFitTable, "表格适应", "将光标所在表格按窗口宽度自动调整", 6, "TableAutoFitWindow"
    ConfigureButton Me.btnNormalizeTerms, "术语下标", "标准化常见科学术语中的下标格式", 7, "Subscript"

    RestoreToolbarPosition
End Sub

Private Sub ConfigureButton(ByVal buttonItem As MSForms.CommandButton, _
                            ByVal captionText As String, _
                            ByVal tipText As String, _
                            ByVal rowIndex As Long, _
                            Optional ByVal imageMsoId As String = "")
    With buttonItem
        .Caption = captionText
        .ControlTipText = captionText & "：" & tipText
        .Tag = captionText
        .Left = BUTTON_LEFT
        .Top = 8 + rowIndex * BUTTON_STEP
        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT
        .PicturePosition = fmPicturePositionCenter
        .TabIndex = rowIndex
        .TabStop = False
        .TakeFocusOnClick = False
        .Font.Name = "微软雅黑"
        .Font.Size = 9
    End With

    If Len(imageMsoId) > 0 Then
        If ApplyOfficeIcon(buttonItem, imageMsoId) Then Exit Sub
    ElseIf HasButtonPicture(buttonItem) Then
        buttonItem.Caption = vbNullString
        Exit Sub
    End If

    ' 如果当前 Office 版本没有对应图标，保留文字作为可用的降级显示。
    buttonItem.Caption = captionText
End Sub

Private Function ApplyOfficeIcon(ByVal buttonItem As MSForms.CommandButton, _
                                 ByVal imageMsoId As String) As Boolean
    On Error GoTo ErrH

    Set buttonItem.Picture = Application.CommandBars.GetImageMso(imageMsoId, 32, 32)
    buttonItem.PicturePosition = fmPicturePositionCenter
    buttonItem.Caption = vbNullString
    ApplyOfficeIcon = True
    Exit Function

ErrH:
    ApplyOfficeIcon = False
End Function

Private Function HasButtonPicture(ByVal buttonItem As MSForms.CommandButton) As Boolean
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

Private Sub btnHeading2_Click()
    RunQuickToolbarAction "style_heading2"
End Sub

Private Sub btnHeading3_Click()
    RunQuickToolbarAction "style_heading3"
End Sub

Private Sub btnBody_Click()
    RunQuickToolbarAction "style_body"
End Sub

Private Sub btnTextBlue_Click()
    RunQuickToolbarAction "text_blue"
End Sub

Private Sub btnPageBreakBefore_Click()
    RunQuickToolbarAction "page_break_before"
End Sub

Private Sub btnAutoFitTable_Click()
    RunQuickToolbarAction "autofit_table"
End Sub

Private Sub btnNormalizeTerms_Click()
    RunQuickToolbarAction "normalize_terms"
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
    Me.Top = Application.Top + 110
    Exit Sub

UseFallback:
    Me.Left = 80
    Me.Top = 80
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
