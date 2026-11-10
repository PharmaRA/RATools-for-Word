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
Private mLabel As MSForms.ComboBox
Private mTitle As MSForms.TextBox
Private mInsertionPoint As Range

Private Sub UserForm_Initialize()
    Dim label As CaptionLabel, hint As MSForms.Label, control As MSForms.Control
    Me.Caption = "生成题注与引用"
    Me.Width = 350
    Me.Height = 206
    Me.Font.Name = "Microsoft YaHei"
    Me.Font.Size = 9

    Set hint = Me.Controls.Add("Forms.Label.1", "lblLabel", True)
    hint.Caption = "题注标签"
    hint.Move 16, 18, 60, 18
    Set mLabel = Me.Controls.Add("Forms.ComboBox.1", "cboLabel", True)
    mLabel.Move 82, 14, 155, 22
    mLabel.Style = fmStyleDropDownList
    mLabel.TabIndex = 0
    For Each label In Application.CaptionLabels
        mLabel.AddItem label.Name
    Next label
    mLabel.Value = GetCaptionPairDefaultLabel()

    Set mNewLabel = Me.Controls.Add("Forms.CommandButton.1", "btnNewLabel", True)
    mNewLabel.Caption = "新建标签…"
    mNewLabel.Move 245, 14, 78, 23
    mNewLabel.TabIndex = 1

    Set hint = Me.Controls.Add("Forms.Label.1", "lblTitle", True)
    hint.Caption = "题注文字"
    hint.Move 16, 56, 60, 18
    Set mTitle = Me.Controls.Add("Forms.TextBox.1", "txtTitle", True)
    mTitle.Move 82, 52, 241, 23
    mTitle.TabIndex = 2

    Set hint = Me.Controls.Add("Forms.Label.1", "lblHint", True)
    hint.Caption = "将在光标处先生成交叉引用，再生成题注。" & vbCrLf & _
                   "生成后可剪切整段题注或引用，分别移到所需位置。" & vbCrLf & _
                   "引用仅显示标签和编号；题注文字可留空后补。"
    hint.Move 16, 88, 307, 46
    hint.ForeColor = RGB(90, 90, 90)

    Set mGenerate = Me.Controls.Add("Forms.CommandButton.1", "btnGenerate", True)
    mGenerate.Caption = "生成"
    mGenerate.Move 161, 145, 76, 25
    mGenerate.Default = True
    mGenerate.TabIndex = 3
    Set mCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    mCancel.Caption = "取消"
    mCancel.Move 247, 145, 76, 25
    mCancel.Cancel = True
    mCancel.TabIndex = 4
    For Each control In Me.Controls
        control.Font.Name = "Microsoft YaHei"
        control.Font.Size = 9
    Next control
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
