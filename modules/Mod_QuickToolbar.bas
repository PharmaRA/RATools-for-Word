Attribute VB_Name = "Mod_QuickToolbar"
Option Explicit

' =============================================
' 快捷悬浮工具栏：窗体生命周期 + 统一动作分发
' v0.8 重构阶段6：原先劈成两半的分发表（前 8 个动作在本模块、
' 后 13 个在 Mod_QuickToolbarActions）合并为单一 Select Case。
' 动作键清单与 scripts/QuickToolbarButtons.psd1 保持一致。
' =============================================

Private Const QUICK_TOOLBAR_TITLE As String = "RATools 快捷工具栏"

' Ribbon 入口：显示或隐藏快捷悬浮窗（customUI14.xml onAction，不可改名）
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
    MsgBox "无法切换快捷工具栏：" & Err.Description, vbExclamation, QUICK_TOOLBAR_TITLE
End Sub

' 无参数入口：用于从宏列表或其他 VBA 过程中调用。
Public Sub ShowQuickToolbar()
    On Error GoTo ErrH
    Load frmQuickToolbar
    frmQuickToolbar.Show vbModeless
    Exit Sub

ErrH:
    MsgBox "无法显示快捷工具栏：" & Err.Description, vbExclamation, QUICK_TOOLBAR_TITLE
End Sub

Public Sub HideQuickToolbar()
    On Error Resume Next
    frmQuickToolbar.SaveToolbarPosition
    frmQuickToolbar.Hide
End Sub

' 按钮统一通过动作键调升本过程（避免窗体复制 Ribbon 业务逻辑）。
' 单一分发表：19 个动作键与悬浮窗按钮一一对应。
Public Sub RunQuickToolbarAction(ByVal actionKey As String)
    On Error GoTo ErrH

    If Documents.Count = 0 Then
        MsgBox "请先打开或新建一个 Word 文档。", vbInformation, QUICK_TOOLBAR_TITLE
        Exit Sub
    End If

    Select Case LCase$(Trim$(actionKey))
        ' --- 编号标题 ---
        Case "style_heading1"
            ApplyRAToolsStyle NumberedHeadingStyle(1)
        Case "style_heading2"
            ApplyRAToolsStyle NumberedHeadingStyle(2)
        Case "style_heading3"
            ApplyRAToolsStyle NumberedHeadingStyle(3)
        Case "style_heading4"
            ApplyRAToolsStyle NumberedHeadingStyle(4)

        ' --- 无编号标题 ---
        Case "style_unnumbered_heading1"
            ApplyRAToolsStyle UnnumberedHeadingStyle(1)
        Case "style_unnumbered_heading2"
            ApplyRAToolsStyle UnnumberedHeadingStyle(2)
        Case "style_unnumbered_heading3"
            ApplyRAToolsStyle UnnumberedHeadingStyle(3)
        Case "style_unnumbered_heading4"
            ApplyRAToolsStyle UnnumberedHeadingStyle(4)

        ' --- 其他样式 ---
        Case "style_body"
            ApplyRAToolsStyle BodyTextStyle()
        Case "style_table_title"
            ApplyRAToolsStyle TableTitleStyle()
        Case "style_figure_title"
            ApplyRAToolsStyle FigureTitleStyle()

        ' --- Word 内置命令 ---
        Case "format_painter"
            ExecuteQuickToolbarMso "FormatPainter"
        Case "paragraph_settings"
            ExecuteQuickToolbarMso "ParagraphDialog"
        Case "insert_cross_reference"
            ExecuteQuickToolbarMso "CrossReferenceInsert"
        Case "update_fields"
            ExecuteQuickToolbarMso "FieldsUpdate"

        ' --- Ribbon 命令复用 ---
        Case "text_blue"
            SetTextBlue Nothing
        Case "page_break_before"
            TogglePageBreakBefore Nothing
        Case "autofit_table"
            AutoFitTableWindow Nothing

        ' --- 批处理宏 ---
        Case "normalize_terms"
            NormalizeScientificTerms
        Case "hyperlinks_fields_blue"
            SetHyperlinksAndFieldsToBlue
        Case "accept_revisions_comments"
            BatchAcceptAndClean
        Case "detect_highlights"
            BatchDetectHighlights

        Case Else
            Err.Raise vbObjectError + 2048, "RunQuickToolbarAction", _
                      "未知的快捷操作：" & actionKey
    End Select
    Exit Sub

ErrH:
    MsgBox "快捷操作执行失败：" & Err.Description, vbExclamation, QUICK_TOOLBAR_TITLE
End Sub

' ====== TEST SUPPORT（仅测试脚本使用，生产路径不调用） ======

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