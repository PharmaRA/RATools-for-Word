VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCaptionReference
   Caption         =   "生成题注与引用"
   ClientHeight    =   3540
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6660
   OleObjectBlob   =   "frmCaptionReference.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "frmCaptionReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mNewLabel As MSForms.CommandButton
Private WithEvents mGenerate As MSForms.CommandButton
Private WithEvents mCancel As MSForms.CommandButton
Private WithEvents mTitle As MSForms.TextBox
Private mLabel As MSForms.ComboBox
Private mLblLabel As MSForms.Label
Private mLblTitle As MSForms.Label
Private mBgTitle As MSForms.Label
Private mBgHint As MSForms.Label
Private mLblHint As MSForms.Label
Private mInsertionPoint As Range
Private mTitleFocused As Boolean
Private mHoverButtons As Collection
Private mFlatNewLabel As clsFlatButton
Private mFlatGenerate As clsFlatButton
Private mFlatCancel As clsFlatButton

Private Sub UserForm_Initialize()
    Dim label As CaptionLabel

    Me.Caption = "生成题注与引用"
    Me.Width = 372
    Me.Height = 228
    Set mHoverButtons = New Collection

    Set mLblLabel = Me.Controls.Add("Forms.Label.1", "lblLabel", True)
    With mLblLabel
        .Caption = "题注标签"
        .Move 16, 18.5, 56, 15
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 9
    End With

    Set mLabel = Me.Controls.Add("Forms.ComboBox.1", "cboLabel", True)
    With mLabel
        .Move 78, 14, 178, 22
        .Style = fmStyleDropDownList
        .TabIndex = 0
    End With
    For Each label In Application.CaptionLabels
        mLabel.AddItem label.Name
    Next label
    mLabel.Value = GetCaptionPairDefaultLabel()

    ' 保留离屏原生按钮以兼容自动化测试（.Value = True）与 Enter / Esc 快捷键
    Set mNewLabel = Me.Controls.Add("Forms.CommandButton.1", "btnNewLabel", True)
    With mNewLabel
        .Caption = "新建标签…"
        .TabStop = False
        .Move -500, -500, 10, 10
    End With

    Set mFlatNewLabel = CreateFlatButton(mHoverButtons, Me, "flatNewLabel", "新建标签…", "secondary", "new_label", Me, 9)
    mFlatNewLabel.MoveTo 264, 14, 82, 22

    Set mLblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle", True)
    With mLblTitle
        .Caption = "题注文字"
        .Move 16, 52.5, 56, 15
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 9
    End With

    Set mBgTitle = Me.Controls.Add("Forms.Label.1", "bgTitle", True)
    mBgTitle.Move 78, 48, 268, 24

    Set mTitle = Me.Controls.Add("Forms.TextBox.1", "txtTitle", True)
    With mTitle
        .Move 85, 53, 254, 14
        .TabIndex = 1
    End With

    Set mBgHint = Me.Controls.Add("Forms.Label.1", "bgHint", True)
    mBgHint.Move 16, 84, 330, 56

    Set mLblHint = Me.Controls.Add("Forms.Label.1", "lblHint", True)
    With mLblHint
        .Caption = "将在光标处先生成交叉引用，再生成题注。" & vbCrLf & _
                   "生成后可剪切整段题注或引用，分别移到所需位置。" & vbCrLf & _
                   "引用仅显示标签和编号；题注文字可留空后补。"
        .Move 26, 91, 310, 44
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
    End With

    Set mGenerate = Me.Controls.Add("Forms.CommandButton.1", "btnGenerate", True)
    With mGenerate
        .Caption = "生成"
        .Default = True
        .TabStop = False
        .Move -500, -500, 10, 10
    End With

    Set mCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With mCancel
        .Caption = "取消"
        .Cancel = True
        .TabStop = False
        .Move -500, -500, 10, 10
    End With

    Set mFlatGenerate = CreateFlatButton(mHoverButtons, Me, "flatGenerate", "生成", "primary", "generate", Me, 9)
    mFlatGenerate.MoveTo 272, 154, 74, 24

    Set mFlatCancel = CreateFlatButton(mHoverButtons, Me, "flatCancel", "取消", "secondary", "cancel", Me, 9)
    mFlatCancel.MoveTo 198, 154, 66, 24

    ApplyTheme
End Sub

Private Sub UserForm_Activate()
    ApplyTheme
End Sub

Public Sub ApplyTheme()
    On Error Resume Next
    ApplyFormBaseTheme Me

    mLblLabel.ForeColor = ThemeTextPrimaryColor()
    mLblTitle.ForeColor = ThemeTextPrimaryColor()
    StyleFlatComboBox mLabel, 9
    StyleFlatInputBackdrop mBgTitle, mTitleFocused
    StyleFlatTextBox mTitle, 9.5
    StyleFlatCard mBgHint
    mLblHint.ForeColor = ThemeTextMutedColor()

    mBgTitle.ZOrder 1
    mTitle.ZOrder 0
    mBgHint.ZOrder 1
    mLblHint.ZOrder 0

    RefreshHoverButtons mHoverButtons
    ApplyFormWindowChrome Me.Caption
    On Error GoTo 0
End Sub

Public Sub OnHoverControlChanged(ByVal activeItem As Object)
    ResetHoverButtons mHoverButtons, activeItem
End Sub

Public Sub OnFlatButtonClick(ByVal actionKey As String)
    Select Case actionKey
        Case "new_label": mNewLabel_Click
        Case "generate": mGenerate_Click
        Case "cancel": mCancel_Click
    End Select
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons mHoverButtons
End Sub

Private Sub mTitle_Enter()
    mTitleFocused = True
    StyleFlatInputBackdrop mBgTitle, True
End Sub

Private Sub mTitle_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    mTitleFocused = False
    StyleFlatInputBackdrop mBgTitle, False
End Sub

Public Sub SetInsertionPoint(ByVal target As Range)
    Set mInsertionPoint = target.Duplicate
End Sub

' 新标签在窗口中暂存，点击生成时才注册到 Word；取消不产生标签副作用。
Public Sub AddLabelChoice(ByVal value As String)
    Dim i As Long
    value = NormalizeCaptionPairLabel(value)
    For i = 0 To mLabel.ListCount - 1
        If StrComp(mLabel.List(i), value, vbTextCompare) = 0 Then
            mLabel.ListIndex = i
            Exit Sub
        End If
    Next i
    mLabel.AddItem value
    mLabel.ListIndex = mLabel.ListCount - 1
End Sub

Private Sub mNewLabel_Click()
    Dim value As String
    value = InputBox("请输入题注标签，例如：附表、Table、Figure。", "新建题注标签")
    If Len(value) = 0 Then Exit Sub
    On Error GoTo Failed
    AddLabelChoice value
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "新建题注标签"
End Sub

Private Sub mGenerate_Click()
    Dim pair As Range
    On Error GoTo Failed
    Set pair = InsertCaptionReferencePair(mInsertionPoint, CStr(mLabel.Value), mTitle.Text)
    RememberCaptionPairLabel CStr(mLabel.Value)
    pair.Collapse wdCollapseEnd
    pair.Select
    Me.Hide
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "生成题注与引用"
End Sub

Private Sub mCancel_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ClearHoverButtons mHoverButtons
End Sub
