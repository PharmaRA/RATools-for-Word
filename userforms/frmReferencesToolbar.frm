VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReferencesToolbar 
   Caption         =   "RATools 图表引用工具"
   ClientHeight    =   6960
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8424.001
   OleObjectBlob   =   "frmReferencesToolbar.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmReferencesToolbar"
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
    Me.Left = CSng(GetSetting("RATools", "ReferencesToolbar", "Left", CStr(Application.Left + 30)))
    Me.Top = CSng(GetSetting("RATools", "ReferencesToolbar", "Top", CStr(Application.Top + 80)))
    On Error GoTo 0
    InitializeMenus
    RefreshDocumentState False
End Sub
Private Sub cboToolbars_Change()
    Dim kind As String
    If mUpdating Or cboToolbars.ListIndex < 1 Then Exit Sub
    Select Case cboToolbars.ListIndex
        Case 1: kind = "Quick"
        Case 2: kind = "Writing"
        Case 3: kind = "Review"
        Case 4: kind = "Final"
        Case 5: kind = "References"
    End Select
    mUpdating = True: cboToolbars.ListIndex = 0: mUpdating = False
    If Len(kind) > 0 Then ToggleWorkflowToolbar kind
End Sub
Private Sub btn0_Click()
    RunWorkflowAction "caption_pair"
End Sub
Private Sub btn1_Click()
    RunWorkflowAction "mso:CaptionInsert"
End Sub
Private Sub btn2_Click()
    RunWorkflowAction "mso:CrossReferenceInsert"
End Sub
Private Sub btn3_Click()
    RunWorkflowAction "restore_caption_labels"
End Sub
Private Sub btn4_Click()
    RunWorkflowAction "style:表标题-F"
End Sub
Private Sub btn5_Click()
    RunWorkflowAction "style:图标题-F"
End Sub
Private Sub btn6_Click()
    RunWorkflowAction "autofit_table"
End Sub
Private Sub cbotable_tools_Change()
    Dim actionKey As String
    If mUpdating Or cbotable_tools.ListIndex < 1 Then Exit Sub
    Select Case cbotable_tools.ListIndex
        Case 2: actionKey = "style:表头左对齐-F"
        Case 3: actionKey = "style:表头居中-F"
        Case 4: actionKey = "style:表头右对齐-F"
        Case 5: actionKey = "style:表格文本左对齐-F"
        Case 6: actionKey = "style:表格文本居中-F"
        Case 7: actionKey = "style:表格文本右对齐-F"
        Case 8: actionKey = "style:表格文本无间距-F"
        Case 9: actionKey = "style:表格注释-F"
        Case 11: actionKey = "mso:TableInsertGallery"
        Case 12: actionKey = "mso:TableBordersMenu"
        Case 13: actionKey = "mso:TableRepeatHeaderRows"
        Case 14: actionKey = "mso:TableSplitTable"
        Case 16: actionKey = "style:图片-F"
        Case 17: actionKey = "mso:TextWrappingInLineWithText"
    End Select
    mUpdating = True: cbotable_tools.ListIndex = 0: mUpdating = False
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
Private Sub cbotable_tools_Enter()
    On Error Resume Next
    Set mMenuRange = Nothing
    If Documents.Count > 0 Then Set mMenuRange = Selection.Range.Duplicate
End Sub
Private Sub cbotable_tools_DropButtonClick()
    cbotable_tools_Enter
End Sub
Private Sub btn8_Click()
    RunWorkflowAction "paste_relink"
End Sub
Private Sub btn9_Click()
    RunWorkflowAction "mso:FieldsUpdate"
End Sub
Private Sub btn10_Click()
    RunWorkflowAction "links_fields_blue"
End Sub
Private Sub btn11_Click()
    RunWorkflowAction "protect_field_format"
End Sub
Private Sub btn12_Click()
    RunWorkflowAction "mso:TableOfFiguresInsert"
End Sub
Private Sub btn13_Click()
    RunWorkflowAction "mso:TableOfFiguresUpdate"
End Sub
Private Sub InitializeMenus()
    mUpdating = True
    cboToolbars.AddItem "工具窗"
    cboToolbars.AddItem "快捷悬浮窗"
    cboToolbars.AddItem "撰写悬浮窗"
    cboToolbars.AddItem "审阅悬浮窗"
    cboToolbars.AddItem "定稿悬浮窗"
    cboToolbars.AddItem "图表引用悬浮窗"
    cboToolbars.ListIndex = 0
    cbotable_tools.AddItem "更多图表工具"
    cbotable_tools.AddItem "— 表格文字 —"
    cbotable_tools.AddItem "表头左对齐"
    cbotable_tools.AddItem "表头居中"
    cbotable_tools.AddItem "表头右对齐"
    cbotable_tools.AddItem "表格正文左对齐"
    cbotable_tools.AddItem "表格正文居中"
    cbotable_tools.AddItem "表格正文右对齐"
    cbotable_tools.AddItem "无间距"
    cbotable_tools.AddItem "表格注释"
    cbotable_tools.AddItem "— 表格结构 —"
    cbotable_tools.AddItem "插入表格"
    cbotable_tools.AddItem "边框"
    cbotable_tools.AddItem "重复标题行"
    cbotable_tools.AddItem "拆分表格"
    cbotable_tools.AddItem "— 图片 —"
    cbotable_tools.AddItem "图片段落样式"
    cbotable_tools.AddItem "嵌入型环绕"
    cbotable_tools.ListIndex = 0
    mUpdating = False
End Sub
Public Sub RefreshDocumentState(ByVal busy As Boolean)
    Dim item As Object, available As Boolean
    On Error GoTo Unavailable
    If cbotable_tools.ListCount = 0 Then InitializeMenus
    If Documents.Count = 0 Then
        lblDocument.Caption = "当前没有打开的文档"
    Else
        lblDocument.Caption = "当前文档：" & ActiveDocument.Name
        available = True
    End If
    GoTo UpdateControls
Unavailable:
    available = False
    lblDocument.Caption = "当前文档不可用"
UpdateControls:
    For Each item In Me.Controls
        If TypeName(item) = "CommandButton" Or TypeName(item) = "ComboBox" Then item.Enabled = Not busy And (available Or item.Name = "cboToolbars")
    Next item
End Sub
Private Sub UserForm_Activate(): RefreshWorkflowToolbarsIfVisible: End Sub
Public Sub SaveToolbarPosition()
    On Error Resume Next
    SaveSetting "RATools", "ReferencesToolbar", "Left", CStr(Me.Left)
    SaveSetting "RATools", "ReferencesToolbar", "Top", CStr(Me.Top)
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer): SaveToolbarPosition: End Sub
Private Sub lblDocument_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub cboToolbars_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
Private Sub lblGroup4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
Private Sub cbotable_tools_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
Private Sub lblGroup12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
Private Sub btn13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EnsureTooltipOnTop Me.Caption
End Sub
