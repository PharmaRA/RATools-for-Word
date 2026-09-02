Attribute VB_Name = "UI_QuickToolbar"
Option Explicit

' =============================================
' 快捷悬浮工具栏：窗体生命周期 + 统一动作分发
' v0.8 重构阶段6：原先劈成两半的分发表（前 8 个动作在本模块、
' 后 13 个在 Mod_QuickToolbarActions）合并为单一 Select Case。
' 动作键清单与 scripts/QuickToolbarButtons.psd1 保持一致。
' =============================================

' 悬浮窗的窗口标题（即窗体 Caption），Win32 按标题定位窗口时使用，
' 必须与 frmQuickToolbar 的 Me.Caption 保持一致。
Private Const QUICK_TOOLBAR_WINDOW_CAPTION As String = "RATools"

' 提示气泡置顶轮询：窗体钉入置顶层后，MSForms 的 ControlTipText 气泡
' （tooltips_class32 窗口）默认不在置顶带，会被窗体压住。用系统定时器周期性
' 把可见的气泡钉回置顶带，窗体隐藏或 Word 退出时停止。


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
    ' Show 之后立刻脱离文档窗口，否则关闭该文档会连带销毁悬浮窗；
    ' 脱离 owner 的同时会把窗体钉进置顶层，避免被 Word 主窗口盖住（见 Mod_Core_Window）。
    DetachQuickToolbarFromDocumentWindow
    ' 显示悬浮窗不应抢走编辑焦点，否则用户回到正文要多点一次。
    ReturnFocusToDocumentWindow
    Exit Sub

ErrH:
    MsgBox "无法显示快捷工具栏：" & Err.Description, vbExclamation, QUICK_TOOLBAR_TITLE
End Sub

Public Sub HideQuickToolbar()
    On Error Resume Next
    frmQuickToolbar.SaveToolbarPosition
    frmQuickToolbar.Hide
End Sub

' 悬浮窗可见期间，用系统定时器周期性把提示气泡钉回置顶带；
' 窗体隐藏或 Word 退出时停止。气泡窗口由 MSForms 按需创建，只有轮询能随时抓住它。




Public Sub RefreshTooltipWatch()
    On Error Resume Next
    If Not IsUserFormWindowVisible(QUICK_TOOLBAR_WINDOW_CAPTION) Then Exit Sub
    EnsureTooltipOnTop QUICK_TOOLBAR_WINDOW_CAPTION
End Sub

' 按钮统一通过动作键调升本过程（避免窗体复制 Ribbon 业务逻辑）。
' 单一分发表：19 个动作键与悬浮窗按钮一一对应。
Public Sub RunQuickToolbarAction(ByVal actionKey As String)
    DispatchQuickToolbarAction actionKey
    ' 点击悬浮窗按钮会把激活状态移到窗体上，之后用户在正文里的第一次点击
    ' 只是把焦点还给 Word，光标要点第二次才落到目标位置。动作结束后主动
    ' 把焦点交还文档窗口，消除这次多余点击。成功与失败路径都要走到。
    ReturnFocusToDocumentWindow
End Sub

Private Sub DispatchQuickToolbarAction(ByVal actionKey As String)
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

        Case "page_break_before"
            TogglePageBreakBefore Nothing
        Case "autofit_table"
            AutoFitTableWindow Nothing

        ' --- 批处理宏 ---

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

' ====== 悬浮窗窗口所有权（脱离单个文档窗口的生命周期） ======
'
' Word 2013+ 是 SDI，modeless UserForm 默认由 Show 时的活动文档窗口 own，
' 关闭该文档会连带销毁窗体。下面几个过程把 owner 置 0 并接管窗体的显示、
' 置顶与卸载；窗口定位统一用标题匹配，避免触碰默认实例导致意外加载窗体。

' 让悬浮窗脱离当前文档窗口；须在 Show 之后调用
Public Function DetachQuickToolbarFromDocumentWindow() As Boolean
    On Error Resume Next
    DetachQuickToolbarFromDocumentWindow = _
        DetachUserFormWindowFromOwner(QUICK_TOOLBAR_WINDOW_CAPTION)
End Function

' 文档切换/激活后再确认一次悬浮窗置顶；未显示时什么都不做。
' 刻意不访问 frmQuickToolbar.Visible：默认实例会在此处被隐式创建。
Public Sub RaiseQuickToolbarIfVisible()
    On Error Resume Next
    If Not IsUserFormWindowVisible(QUICK_TOOLBAR_WINDOW_CAPTION) Then Exit Sub
    KeepUserFormWindowOnTop QUICK_TOOLBAR_WINDOW_CAPTION
End Sub

' 把激活状态交还 Word 文档窗口，并确认悬浮窗仍在最前。
'
' 点击 modeless UserForm 会激活窗体自身的顶层窗口，Word 随之失活；用户回到
' 正文的第一次点击只是重新激活 Word，光标要到第二次点击才落到目标位置。
' 每个动作结束后调用本过程，把那次多余点击省掉。
'
' 用 Application.Activate 而非 SetForegroundWindow：前者是文档模型内的受支持
' 做法，且不受前台窗口切换限制的影响。无文档打开时静默返回。
'
' 激活 Word 会把文档窗口抬到 z 序最前，owner 为 0 的悬浮窗只靠置顶层留在前面，
' 所以交还焦点后再确认一次置顶。少了这一步，点完一个按钮悬浮窗就被 Word 盖到
' 后面，而工具窗没有任务栏按钮，用户只能从 Ribbon 关掉再打开才能接着点。
Public Sub ReturnFocusToDocumentWindow()
    On Error Resume Next
    If Documents.Count = 0 Then Exit Sub
    Application.Activate
    RaiseQuickToolbarIfVisible
End Sub

' Word 退出前卸载悬浮窗：owner 置 0 后它不再随文档窗口销毁，
' 必须显式卸载，否则位置来不及存盘。
Public Sub UnloadQuickToolbarOnQuit()
    On Error Resume Next
    If FindUserFormWindow(QUICK_TOOLBAR_WINDOW_CAPTION) = 0 Then Exit Sub
    frmQuickToolbar.SaveToolbarPosition
    Unload frmQuickToolbar
End Sub
