VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWritingToolbar 
   Caption         =   "RATools 撰写工具"
   ClientHeight    =   6720
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7620
   OleObjectBlob   =   "frmWritingToolbar.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmWritingToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mUpdating As Boolean
Private mMenuRange As Range
Private Sub UserForm_Initialize()
    Dim item As Object
    For Each item In Me.Controls: item.Font.Name = "Microsoft YaHei": item.Font.Size = 9: Next item
    On Error Resume Next
    Me.Left = CSng(GetSetting("RATools", "WritingToolbar", "Left", CStr(Application.Left + 30)))
    Me.Top = CSng(GetSetting("RATools", "WritingToolbar", "Top", CStr(Application.Top + 80)))
    On Error GoTo 0
    InitializeMenus
    RefreshDocumentState False
End Sub
Private Sub btn0_Click()
    RunWorkflowAction "style:标题1-F"
End Sub
Private Sub btn1_Click()
    RunWorkflowAction "style:标题2-F"
End Sub
Private Sub btn2_Click()
    RunWorkflowAction "style:标题3-F"
End Sub
Private Sub btn3_Click()
    RunWorkflowAction "style:标题4-F"
End Sub
Private Sub btn4_Click()
    RunWorkflowAction "style:无编号标题1-F"
End Sub
Private Sub btn5_Click()
    RunWorkflowAction "style:无编号标题2-F"
End Sub
Private Sub btn6_Click()
    RunWorkflowAction "style:无编号标题3-F"
End Sub
Private Sub btn7_Click()
    RunWorkflowAction "style:无编号标题4-F"
End Sub
Private Sub btn8_Click()
    RunWorkflowAction "style:正文-F"
End Sub
Private Sub btn9_Click()
    RunWorkflowAction "style:正文无缩进-F"
End Sub
Private Sub btn10_Click()
    RunWorkflowAction "style:正文无间距-F"
End Sub
Private Sub cbostyles_Change()
    Dim actionKey As String
    If mUpdating Or cbostyles.ListIndex < 1 Then Exit Sub
    Select Case cbostyles.ListIndex
        Case 2: actionKey = "style:标题5-F"
        Case 3: actionKey = "style:标题6-F"
        Case 4: actionKey = "style:标题7-F"
        Case 5: actionKey = "style:标题8-F"
        Case 6: actionKey = "style:标题9-F"
        Case 8: actionKey = "style:无编号标题5-F"
        Case 9: actionKey = "style:无编号标题6-F"
        Case 10: actionKey = "style:无编号标题7-F"
        Case 11: actionKey = "style:无编号标题8-F"
        Case 12: actionKey = "style:无编号标题9-F"
        Case 14: actionKey = "style:标题居中-F"
        Case 15: actionKey = "style:标题左对齐-F"
        Case 16: actionKey = "style:脚注-F"
        Case 17: actionKey = "style:参考文献列表-F"
        Case 18: actionKey = "style_pane"
    End Select
    mUpdating = True: cbostyles.ListIndex = 0: mUpdating = False
    On Error Resume Next
    If Not mMenuRange Is Nothing Then
        If Documents.Count > 0 Then
            If mMenuRange.Document Is ActiveDocument Then mMenuRange.Select
        End If
    End If
    Set mMenuRange = Nothing
    On Error GoTo 0
    If Len(actionKey) > 0 Then RunWorkflowAction actionKey
End Sub
Private Sub cbostyles_Enter()
    On Error Resume Next
    Set mMenuRange = Nothing
    If Documents.Count > 0 Then Set mMenuRange = Selection.Range.Duplicate
End Sub
Private Sub cbostyles_DropButtonClick()
    cbostyles_Enter
End Sub
Private Sub btn12_Click()
    RunWorkflowAction "style:编号列表-F"
End Sub
Private Sub btn13_Click()
    RunWorkflowAction "style:项目符号列表-F"
End Sub
Private Sub btn14_Click()
    RunWorkflowAction "mso:FootnoteInsert"
End Sub
Private Sub cbotext_Change()
    Dim actionKey As String
    If mUpdating Or cbotext.ListIndex < 1 Then Exit Sub
    Select Case cbotext.ListIndex
        Case 1: actionKey = "mso:Superscript"
        Case 2: actionKey = "mso:Subscript"
        Case 3: actionKey = "normalize_terms"
        Case 4: actionKey = "extract_abbreviations"
    End Select
    mUpdating = True: cbotext.ListIndex = 0: mUpdating = False
    On Error Resume Next
    If Not mMenuRange Is Nothing Then
        If Documents.Count > 0 Then
            If mMenuRange.Document Is ActiveDocument Then mMenuRange.Select
        End If
    End If
    Set mMenuRange = Nothing
    On Error GoTo 0
    If Len(actionKey) > 0 Then RunWorkflowAction actionKey
End Sub
Private Sub cbotext_Enter()
    On Error Resume Next
    Set mMenuRange = Nothing
    If Documents.Count > 0 Then Set mMenuRange = Selection.Range.Duplicate
End Sub
Private Sub cbotext_DropButtonClick()
    cbotext_Enter
End Sub
Private Sub btn16_Click()
    RunWorkflowAction "mso:FormatPainter"
End Sub
Private Sub btn17_Click()
    RunWorkflowAction "mso:ClearFormatting"
End Sub
Private Sub btn18_Click()
    RunWorkflowAction "mso:ParagraphDialog"
End Sub
Private Sub btn19_Click()
    RunWorkflowAction "page_break_before"
End Sub
Private Sub btn20_Click()
    RunWorkflowAction "autofit_table"
End Sub
Private Sub btn21_Click()
    RunWorkflowAction "text_blue"
End Sub
Private Sub btn22_Click()
    RunWorkflowAction "mso:PageSetupDialog"
End Sub
Private Sub cbolayout_Change()
    Dim actionKey As String
    If mUpdating Or cbolayout.ListIndex < 1 Then Exit Sub
    Select Case cbolayout.ListIndex
        Case 2: actionKey = "mso:PageBreakInsertWord"
        Case 3: actionKey = "mso:SectionBreakInsert"
        Case 4: actionKey = "style:页眉-F"
        Case 5: actionKey = "style:页脚-F"
        Case 7: actionKey = "autofit_all_tables"
        Case 8: actionKey = "link_sections"
        Case 9: actionKey = "batch_margins"
    End Select
    mUpdating = True: cbolayout.ListIndex = 0: mUpdating = False
    On Error Resume Next
    If Not mMenuRange Is Nothing Then
        If Documents.Count > 0 Then
            If mMenuRange.Document Is ActiveDocument Then mMenuRange.Select
        End If
    End If
    Set mMenuRange = Nothing
    On Error GoTo 0
    If Len(actionKey) > 0 Then RunWorkflowAction actionKey
End Sub
Private Sub cbolayout_Enter()
    On Error Resume Next
    Set mMenuRange = Nothing
    If Documents.Count > 0 Then Set mMenuRange = Selection.Range.Duplicate
End Sub
Private Sub cbolayout_DropButtonClick()
    cbolayout_Enter
End Sub
Private Sub InitializeMenus()
    mUpdating = True
    cbostyles.AddItem "更多样式"
    cbostyles.AddItem "— 编号标题 —"
    cbostyles.AddItem "标题 5"
    cbostyles.AddItem "标题 6"
    cbostyles.AddItem "标题 7"
    cbostyles.AddItem "标题 8"
    cbostyles.AddItem "标题 9"
    cbostyles.AddItem "— 无编号标题 —"
    cbostyles.AddItem "无编号 5"
    cbostyles.AddItem "无编号 6"
    cbostyles.AddItem "无编号 7"
    cbostyles.AddItem "无编号 8"
    cbostyles.AddItem "无编号 9"
    cbostyles.AddItem "— 其他样式 —"
    cbostyles.AddItem "标题居中"
    cbostyles.AddItem "标题左对齐"
    cbostyles.AddItem "脚注样式"
    cbostyles.AddItem "参考文献列表"
    cbostyles.AddItem "样式窗格"
    cbostyles.ListIndex = 0
    cbotext.AddItem "更多文字"
    cbotext.AddItem "上标"
    cbotext.AddItem "下标"
    cbotext.AddItem "标准化科学术语下标"
    cbotext.AddItem "提取缩略语"
    cbotext.ListIndex = 0
    cbolayout.AddItem "更多排版"
    cbolayout.AddItem "— 局部操作 —"
    cbolayout.AddItem "插入分页符"
    cbolayout.AddItem "插入分节符"
    cbolayout.AddItem "页眉样式"
    cbolayout.AddItem "页脚样式"
    cbolayout.AddItem "— 整篇处理 —"
    cbolayout.AddItem "当前文档表格统一适宽"
    cbolayout.AddItem "页眉页脚链接到前一节"
    cbolayout.AddItem "批量设置页边距〔各节四边 2.54 cm〕"
    cbolayout.ListIndex = 0
    mUpdating = False
End Sub
Public Sub RefreshDocumentState(ByVal busy As Boolean)
    Dim item As Object, available As Boolean, value As Long
    On Error GoTo Unavailable
    If cbostyles.ListCount = 0 Then InitializeMenus
    If Documents.Count = 0 Then lblDocument.Caption = "当前没有打开的文档": GoTo UpdateControls
    lblDocument.Caption = "当前文档：" & ActiveDocument.Name
    available = True
    value = Selection.ParagraphFormat.PageBreakBefore
    btn19.Caption = "段前分页：" & IIf(value = wdUndefined, "混合", IIf(value = 0, "关", "开"))
    GoTo UpdateControls
Unavailable:
    available = False: lblDocument.Caption = "当前文档不可用"
UpdateControls:
    For Each item In Me.Controls
        If TypeName(item) = "CommandButton" Or TypeName(item) = "ComboBox" Then item.Enabled = available And Not busy
    Next item
End Sub
Private Sub UserForm_Activate(): RefreshWorkflowToolbarsIfVisible: End Sub
Public Sub SaveToolbarPosition()
    On Error Resume Next
    SaveSetting "RATools", "WritingToolbar", "Left", CStr(Me.Left)
    SaveSetting "RATools", "WritingToolbar", "Top", CStr(Me.Top)
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer): SaveToolbarPosition: End Sub
Private Sub lblDocument_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
Private Sub cbostyles_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblGroup12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub cbotext_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblGroup16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub lblGroup20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub cbolayout_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
