VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuickToolbar
   Caption         =   "RATools"
   ClientHeight    =   4700
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   1600
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
Private Const TOOLBAR_WIDTH As Single = 82
Private Const TOOLBAR_HEIGHT As Single = 256
Private Const BUTTON_TOP As Single = 4
Private Const BUTTON_WIDTH As Single = 28
Private Const BUTTON_HEIGHT As Single = 25
Private Const BUTTON_STEP As Single = 27

' 图标直接复用 Ribbon 资源：H1/H2/H3、BodyText、SetTextBlue、PageBreakBefore，
' 表格适应和术语下标分别沿用 Ribbon 的 TableAutoFitWindow、Subscript 图标。

Private Sub UserForm_Initialize()
    Me.Caption = "RATools"
    Me.Width = TOOLBAR_WIDTH
    Me.Height = TOOLBAR_HEIGHT
    Me.BackColor = RGB(247, 248, 250)

    ConfigureButton Me.btnHeading1, "标题 1", "应用“标题1-F”样式", 0, "H1"
    ConfigureButton Me.btnHeading2, "标题 2", "应用“标题2-F”样式", 1, "H2"
    ConfigureButton Me.btnHeading3, "标题 3", "应用“标题3-F”样式", 2, "H3"
    ConfigureButton Me.btnBody, "正文", "应用“正文-F”样式", 3, "BodyText"
    ConfigureButton Me.btnTextBlue, "设为蓝色", "将选中文字设置为标准蓝色", 4, "SetTextBlue"
    ConfigureButton Me.btnPageBreakBefore, "段前分页", "切换选中段落的段前分页属性", 5, "PageBreakBefore"
    ConfigureButton Me.btnAutoFitTable, "表格适应", "将光标所在表格按窗口宽度自动调整", 6, "imageMso:TableAutoFitWindow"
    ConfigureButton Me.btnNormalizeTerms, "术语下标", "标准化常见科学术语中的下标格式", 7, "idMso:Subscript"

    RestoreToolbarPosition
End Sub

Private Sub ConfigureButton(ByVal buttonItem As MSForms.Label, _
                            ByVal captionText As String, _
                            ByVal tipText As String, _
                            ByVal rowIndex As Long, _
                            ByVal iconSource As String)
    With buttonItem
        .Caption = captionText
        .ControlTipText = captionText & "：" & tipText
        .Tag = iconSource
        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT
        .Left = (Me.InsideWidth - .Width) / 2
        .Top = BUTTON_TOP + rowIndex * BUTTON_STEP
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(221, 226, 232)
        .PicturePosition = fmPicturePositionCenter
        .TextAlign = fmTextAlignCenter
        .TabIndex = rowIndex
        .Font.Name = "微软雅黑"
        .Font.Size = 7
    End With

    If HasButtonPicture(buttonItem) Then
        buttonItem.Caption = vbNullString
        Exit Sub
    End If

    ' 嵌入图标异常时保留文字，确保功能仍然可识别和使用。
    buttonItem.Caption = captionText
End Sub

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
