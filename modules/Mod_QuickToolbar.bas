Attribute VB_Name = "Mod_QuickToolbar"
Option Explicit

Private Const QUICK_TOOLBAR_TITLE As String = "RATools 快捷工具栏"

' Ribbon 入口：显示或隐藏快捷悬浮窗。
Public Sub ToggleQuickToolbar(ByVal control As IRibbonControl)
    On Error GoTo ErrH

    If frmQuickToolbar.Visible Then
        frmQuickToolbar.SaveToolbarPosition
        frmQuickToolbar.Hide
    Else
        ShowQuickToolbar
    End If
    Exit Sub

ErrH:
    MsgBox "无法切换快捷悬浮窗：" & Err.Description, vbExclamation, QUICK_TOOLBAR_TITLE
End Sub

' 无参数入口，便于从宏列表或其他 VBA 代码中调用。
Public Sub ShowQuickToolbar()
    On Error GoTo ErrH
    Load frmQuickToolbar
    frmQuickToolbar.Show vbModeless
    Exit Sub

ErrH:
    MsgBox "无法显示快捷悬浮窗：" & Err.Description, vbExclamation, QUICK_TOOLBAR_TITLE
End Sub

Public Function GetQuickToolbarButtonCount() As Long
    On Error GoTo ErrH
    Load frmQuickToolbar
    GetQuickToolbarButtonCount = frmQuickToolbar.Controls.Count
    Exit Function

ErrH:
    GetQuickToolbarButtonCount = 0
End Function

Public Function ReleaseQuickToolbarForTest() As Boolean
    On Error GoTo ErrH
    Unload frmQuickToolbar
    ReleaseQuickToolbarForTest = True
    Exit Function

ErrH:
    ReleaseQuickToolbarForTest = False
End Function

Public Sub HideQuickToolbar()
    On Error Resume Next
    frmQuickToolbar.SaveToolbarPosition
    frmQuickToolbar.Hide
End Sub

' 按钮统一通过动作键进入这里，避免窗体复制 Ribbon 业务逻辑。
Public Sub RunQuickToolbarAction(ByVal actionKey As String)
    On Error GoTo ErrH

    If Documents.Count = 0 Then
        MsgBox "请先打开或新建一个 Word 文档。", vbInformation, QUICK_TOOLBAR_TITLE
        Exit Sub
    End If

    Select Case LCase$(Trim$(actionKey))
        Case "style_heading1"
            ApplyRAToolsStyle NumberedHeadingStyle(1)
        Case "style_heading2"
            ApplyRAToolsStyle NumberedHeadingStyle(2)
        Case "style_heading3"
            ApplyRAToolsStyle NumberedHeadingStyle(3)
        Case "style_body"
            ApplyRAToolsStyle BodyTextStyle()
        Case "text_blue"
            SetTextBlue Nothing
        Case "page_break_before"
            TogglePageBreakBefore Nothing
        Case "autofit_table"
            AutoFitTableWindow Nothing
        Case "normalize_terms"
            NormalizeScientificTerms
        Case Else
            If TryRunExpandedQuickToolbarAction(actionKey) Then Exit Sub
            Err.Raise vbObjectError + 2048, "RunQuickToolbarAction", _
                      "未知的快捷操作：" & actionKey
    End Select
    Exit Sub

ErrH:
    MsgBox "快捷操作执行失败：" & Err.Description, vbExclamation, QUICK_TOOLBAR_TITLE
End Sub
