VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMacroList 
   Caption         =   "UserForm1"
   ClientHeight    =   5436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7800
   OleObjectBlob   =   "frmMacroList.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmMacroList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' Windows API 声明 (实现窗体拖拽调整大小)
' ==============================================================================
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000

Private Const MACRO_ROW_HEIGHT As Single = 24
Private Const MACRO_ROW_STEP As Single = 26

Private m_AllMacros As Variant
Private m_IsInitialized As Boolean
Private m_SearchFocused As Boolean
Private m_SyncingSelection As Boolean
Private m_HoverButtons As Collection
Private m_MacroRows As Collection
Private m_CurrentEngName As String

' 扁平化界面元素（隐藏原生 ListBox 与 3D 按钮，仅保留其接口与键盘语义）
Private m_BgSearch As MSForms.Label
Private WithEvents m_LblSearchIcon As MSForms.Label
Private WithEvents m_LblPlaceholder As MSForms.Label
Private m_FlatClear As clsFlatButton
Private m_FlatRefresh As clsFlatButton
Private m_FlatClose As clsFlatButton
Private m_FlatRun As clsFlatButton

Private m_BgListHeader As MSForms.Label
Private m_LblListHeader As MSForms.Label
Private m_LblListCount As MSForms.Label
Private WithEvents m_FrmList As MSForms.Frame

Private m_BgDetailCard As MSForms.Label
Private m_BgCategory As MSForms.Label
Private m_LblCategory As MSForms.Label
Private m_LblDetailTitle As MSForms.Label
Private m_LblMacroCaption As MSForms.Label
Private m_BgMacroBadge As MSForms.Label
Private m_LblMacroBadge As MSForms.Label
Private m_DivDetail As MSForms.Label
Private m_LblDescCaption As MSForms.Label
Private m_LblHint1 As MSForms.Label
Private m_BgKeyEnter As MSForms.Label
Private m_LblKeyEnter As MSForms.Label
Private m_LblHint2 As MSForms.Label
Private m_BgKeyEsc As MSForms.Label
Private m_LblKeyEsc As MSForms.Label
Private m_LblHint3 As MSForms.Label

Private Sub UserForm_Initialize()
    Me.Caption = "宏管理器"
    Me.Width = 548
    Me.Height = 380

    Set m_HoverButtons = New Collection
    Set m_MacroRows = New Collection

    EnsureExtraControls
    ConfigureStaticControls
    RegisterFlatControls

    Me.lstMacros.ColumnCount = 3
    m_IsInitialized = True

    ApplyTheme
    LayoutControls
    LoadDataToMemory
    MakeFormResizable
End Sub

Private Sub UserForm_Activate()
    MakeFormResizable
    ApplyTheme
End Sub

Private Sub EnsureExtraControls()
    Set m_BgSearch = EnsureLabel("bgSearch")
    Set m_LblSearchIcon = EnsureLabel("lblSearchIcon")
    Set m_LblPlaceholder = EnsureLabel("lblSearchPlaceholder")

    Set m_BgListHeader = EnsureLabel("bgListHeader")
    Set m_LblListHeader = EnsureLabel("lblListHeader")
    Set m_LblListCount = EnsureLabel("lblListCount")
    Set m_FrmList = EnsureFrame("frmMacroRows")

    Set m_BgDetailCard = EnsureLabel("bgDetailCard")
    Set m_BgCategory = EnsureLabel("bgCategory")
    Set m_LblCategory = EnsureLabel("lblCategory")
    Set m_LblDetailTitle = EnsureLabel("lblDetailTitle")
    Set m_LblMacroCaption = EnsureLabel("lblMacroCaption")
    Set m_BgMacroBadge = EnsureLabel("bgMacroBadge")
    Set m_LblMacroBadge = EnsureLabel("lblMacroBadge")
    Set m_DivDetail = EnsureLabel("divDetail")
    Set m_LblDescCaption = EnsureLabel("lblDescCaption")
    Set m_LblHint1 = EnsureLabel("lblHint1")
    Set m_BgKeyEnter = EnsureLabel("bgKeyEnter")
    Set m_LblKeyEnter = EnsureLabel("lblKeyEnter")
    Set m_LblHint2 = EnsureLabel("lblHint2")
    Set m_BgKeyEsc = EnsureLabel("bgKeyEsc")
    Set m_LblKeyEsc = EnsureLabel("lblKeyEsc")
    Set m_LblHint3 = EnsureLabel("lblHint3")
End Sub

Private Function EnsureLabel(ByVal ctrlName As String) As MSForms.Label
    On Error Resume Next
    Set EnsureLabel = Me.Controls(ctrlName)
    On Error GoTo 0
    If EnsureLabel Is Nothing Then
        Set EnsureLabel = Me.Controls.Add("Forms.Label.1", ctrlName, True)
    End If
End Function

Private Function EnsureFrame(ByVal ctrlName As String) As MSForms.Frame
    On Error Resume Next
    Set EnsureFrame = Me.Controls(ctrlName)
    On Error GoTo 0
    If EnsureFrame Is Nothing Then
        Set EnsureFrame = Me.Controls.Add("Forms.Frame.1", ctrlName, True)
    End If
End Function

Private Sub ConfigureStaticControls()
    On Error Resume Next
    With Me.Label_1
        .Visible = False
        .Move -500, -500, 10, 10
    End With

    With m_LblSearchIcon
        .Caption = ChrW$(&HE721)
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .Font.Name = ThemeIconFontName()
        .Font.Size = 9
        .ControlTipText = "搜索功能"
    End With

    With Me.txtSearch
        .ControlTipText = "输入中文名称、宏代码名或说明关键词过滤"
        .TabIndex = 0
    End With

    With m_LblPlaceholder
        .Caption = "搜索功能名称、宏代码名或说明关键词..."
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 9
        .WordWrap = False
    End With

    ' 保留原生控件与测试契约（txtSearch / lstMacros / btnRun / btnClose），视觉由自定义层接管
    With Me.lstMacros
        .Visible = False
        .TabStop = False
    End With

    With Me.btnRun
        .Caption = "运行"
        .Default = True
        .TabStop = False
        .Move -500, -500, 10, 10
    End With

    With Me.btnClose
        .Caption = "关闭"
        .Cancel = True
        .TabStop = False
        .Move -500, -500, 10, 10
    End With

    With m_LblListHeader
        .Caption = "功能列表"
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
        .Font.Bold = True
    End With

    With m_LblListCount
        .Caption = "共 0 项"
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignRight
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
    End With

    With m_FrmList
        .Caption = vbNullString
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .ScrollBars = fmScrollBarsNone
    End With

    With m_BgCategory
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With

    With m_LblCategory
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
        .Font.Bold = True
    End With

    With m_LblDetailTitle
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 11
        .Font.Bold = True
        .WordWrap = True
    End With

    With m_BgMacroBadge
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With

    With m_LblMacroCaption
        .Caption = "入口宏名称 (Application.Run)"
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
    End With

    With m_LblMacroBadge
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .Font.Name = ThemeCodeFontName()
        .Font.Size = 8.5
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
    End With

    With m_LblDescCaption
        .Caption = "功能说明"
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
        .Font.Bold = True
    End With

    With m_LblHint1
        .Caption = "提示：双击列表项或按"
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
    End With

    With m_BgKeyEnter
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With

    With m_LblKeyEnter
        .Caption = "Enter"
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .Font.Name = ThemeCodeFontName()
        .Font.Size = 8
    End With

    With m_LblHint2
        .Caption = "运行，"
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
    End With

    With m_BgKeyEsc
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With

    With m_LblKeyEsc
        .Caption = "Esc"
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .Font.Name = ThemeCodeFontName()
        .Font.Size = 8
    End With

    With m_LblHint3
        .Caption = "关闭"
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
    End With
    On Error GoTo 0
End Sub

Private Sub RegisterFlatControls()
    Set m_FlatClear = CreateFlatButton(m_HoverButtons, Me, "flatClear", vbNullString, "ghost", "clear", Me, 8.5)
    m_FlatClear.SetIcon ChrW$(&HE8BB), ThemeIconFontName(), 7
    m_FlatClear.SetToolTip "清空搜索框"
    m_FlatClear.SetVisible False

    Set m_FlatRefresh = CreateFlatButton(m_HoverButtons, Me, "flatRefresh", "刷新", "secondary", "refresh", Me, 9)
    m_FlatRefresh.SetIcon ChrW$(&HE72C), ThemeIconFontName(), 8.5
    m_FlatRefresh.SetToolTip "重新载入宏列表并清空筛选"

    Set m_FlatClose = CreateFlatButton(m_HoverButtons, Me, "flatClose", "关闭", "secondary", "close", Me, 9)

    Set m_FlatRun = CreateFlatButton(m_HoverButtons, Me, "flatRun", "运行功能", "primary", "run", Me, 9)
    m_FlatRun.SetIcon ChrW$(&HF5B0), ThemeIconFontName(), 7.5
End Sub

Public Sub ApplyTheme()
    Dim rowItem As clsMacroListRow

    On Error Resume Next
    ApplyFormBaseTheme Me

    m_LblSearchIcon.ForeColor = ThemeTextMutedColor()
    StyleFlatInputBackdrop m_BgSearch, m_SearchFocused
    StyleFlatTextBox Me.txtSearch, 9.5
    m_LblPlaceholder.ForeColor = ThemeTextMutedColor()

    With m_BgListHeader
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BackColor = ThemeCanvasColor()
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .BorderColor = ThemeBorderColor()
    End With
    m_LblListHeader.ForeColor = ThemeTextMutedColor()
    m_LblListCount.ForeColor = ThemeTextMutedColor()
    StyleFlatCard m_FrmList

    If Not m_MacroRows Is Nothing Then
        For Each rowItem In m_MacroRows
            rowItem.ApplyVisualState
        Next rowItem
    End If

    StyleFlatCard m_BgDetailCard
    With m_BgCategory
        .BackColor = ThemeHoverFillColor()
        .BorderColor = ThemeBorderColor()
    End With
    m_LblCategory.ForeColor = ThemeAccentColor()
    m_LblDetailTitle.ForeColor = ThemeTextPrimaryColor()
    m_LblMacroCaption.ForeColor = ThemeTextMutedColor()
    With m_BgMacroBadge
        .BackColor = ThemeCanvasColor()
        .BorderColor = ThemeBorderColor()
    End With
    m_LblMacroBadge.ForeColor = ThemeAccentColor()
    StyleDividerLine m_DivDetail
    m_LblDescCaption.ForeColor = ThemeTextMutedColor()

    With m_BgKeyEnter
        .BackColor = ThemeCardColor()
        .BorderColor = ThemeBorderColor()
    End With
    With m_BgKeyEsc
        .BackColor = ThemeCardColor()
        .BorderColor = ThemeBorderColor()
    End With
    m_LblHint1.ForeColor = ThemeTextMutedColor()
    m_LblKeyEnter.ForeColor = ThemeTextMutedColor()
    m_LblHint2.ForeColor = ThemeTextMutedColor()
    m_LblKeyEsc.ForeColor = ThemeTextMutedColor()
    m_LblHint3.ForeColor = ThemeTextMutedColor()

    StyleFlatTextBox Me.lblDetail, 9.5
    With Me.lblDetail
        .MultiLine = True
        .WordWrap = True
        .Locked = True
        .TabStop = False
        .ScrollBars = fmScrollBarsVertical
        .ForeColor = ThemeTextMutedColor()
        .BackColor = ThemeCardColor()
    End With

    RefreshHoverButtons m_HoverButtons
    ApplyFormWindowChrome Me.Caption
    On Error GoTo 0
End Sub

Public Sub OnHoverControlChanged(ByVal activeItem As Object)
    ResetHoverButtons m_HoverButtons, activeItem
    ResetMacroRowHover
End Sub

Public Sub OnFlatButtonClick(ByVal actionKey As String)
    Select Case actionKey
        Case "clear"
            m_BtnClearSearch_Click
        Case "refresh"
            m_BtnRefresh_Click
        Case "close"
            btnClose_Click
        Case "run"
            btnRun_Click
    End Select
End Sub

Public Sub OnMacroRowHover(ByVal activeRow As Object)
    Dim rowItem As clsMacroListRow

    ResetHoverButtons m_HoverButtons
    If m_MacroRows Is Nothing Then Exit Sub
    For Each rowItem In m_MacroRows
        If Not (rowItem Is activeRow) Then rowItem.ResetHover
    Next rowItem
End Sub

Public Sub OnMacroRowClick(ByVal idx As Long)
    SelectRowIndex idx
End Sub

Public Sub OnMacroRowDblClick(ByVal idx As Long)
    SelectRowIndex idx
    btnRun_Click
End Sub

Private Sub ResetMacroRowHover()
    Dim rowItem As clsMacroListRow

    If m_MacroRows Is Nothing Then Exit Sub
    For Each rowItem In m_MacroRows
        rowItem.ResetHover
    Next rowItem
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    ResetMacroRowHover
End Sub

Private Sub m_FrmList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    ResetMacroRowHover
End Sub

Private Sub MakeFormResizable()
    Dim hWnd As LongPtr
    Dim iStyle As Long

    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    If hWnd = 0 Then Exit Sub
    iStyle = GetWindowLong(hWnd, GWL_STYLE)
    iStyle = iStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    SetWindowLong hWnd, GWL_STYLE, iStyle
End Sub

Private Sub UserForm_Resize()
    If Not m_IsInitialized Then Exit Sub
    LayoutControls
End Sub

Private Sub LayoutControls()
    Dim pad As Single
    Dim w As Single
    Dim h As Single
    Dim searchLeft As Single
    Dim searchW As Single
    Dim refreshW As Single
    Dim contentTop As Single
    Dim contentH As Single
    Dim headerH As Single
    Dim leftColW As Single
    Dim rightColL As Single
    Dim rightColW As Single
    Dim innerL As Single
    Dim innerW As Single
    Dim descH As Single
    Dim barTop As Single
    Dim hintW As Single
    Dim runW As Single
    Dim closeW As Single
    Dim keyX As Single

    On Error Resume Next
    pad = 14
    w = Me.InsideWidth
    h = Me.InsideHeight
    If w < 380 Then w = 380
    If h < 250 Then h = 250

    ' 顶部搜索栏（框内左放大镜 + 占位符对齐列表左边线）
    refreshW = 62
    m_FlatRefresh.MoveTo w - pad - refreshW, pad, refreshW, 24

    searchLeft = pad
    searchW = w - pad - refreshW - 8 - searchLeft
    If searchW < 120 Then searchW = 120
    m_BgSearch.Move searchLeft, pad, searchW, 24
    m_LblSearchIcon.Move searchLeft + 8, pad + 6, 12, 12
    m_FlatClear.MoveTo searchLeft + searchW - 22, pad + 3, 18, 18
    Me.txtSearch.Move searchLeft + 25, pad + 5.5, searchW - 50, 14
    m_LblPlaceholder.Move searchLeft + 25, pad + 5.5, searchW - 50, 14

    ' 双栏内容区（左列表与右卡片底边完全齐平）
    contentTop = pad + 34
    contentH = h - contentTop - 44
    If contentH < 130 Then contentH = 130

    leftColW = Int((w - pad * 2 - 12) * 0.43)
    If leftColW < 160 Then leftColW = 160
    rightColL = pad + leftColW + 12
    rightColW = w - rightColL - pad
    If rightColW < 180 Then rightColW = 180

    headerH = 24
    m_BgListHeader.Move pad, contentTop, leftColW, headerH
    m_LblListHeader.Move pad + 10, contentTop + 5.5, 64, 13
    m_LblListCount.Move pad + leftColW - 88, contentTop + 5.5, 78, 13
    m_FrmList.Move pad, contentTop + headerH - 0.75, leftColW, contentH - headerH + 0.75

    m_BgDetailCard.Move rightColL, contentTop, rightColW, contentH
    innerL = rightColL + 16
    innerW = rightColW - 32
    If innerW < 80 Then innerW = 80

    m_BgCategory.Move innerL, contentTop + 12, 62, 18
    m_LblCategory.Move innerL, contentTop + 14.5, 62, 13
    m_LblDetailTitle.Move innerL, contentTop + 35, innerW, 20
    m_DivDetail.Move innerL, contentTop + 59, innerW, 0.75
    m_LblMacroCaption.Move innerL, contentTop + 68, innerW, 13
    m_BgMacroBadge.Move innerL, contentTop + 85, innerW, 19
    m_LblMacroBadge.Move innerL + 7, contentTop + 88, innerW - 14, 13
    m_LblDescCaption.Move innerL, contentTop + 112, innerW, 14

    descH = contentH - 135
    If descH < 40 Then descH = 40
    Me.lblDetail.Move innerL - 2, contentTop + 129, innerW + 2, descH

    ' 底部操作栏（带 1px 键帽提示）
    barTop = h - 34
    runW = 86
    closeW = 64
    m_FlatRun.MoveTo w - pad - runW, barTop, runW, 24
    m_FlatClose.MoveTo w - pad - runW - 8 - closeW, barTop, closeW, 24
    hintW = w - pad - runW - 8 - closeW - pad - 10
    If hintW < 60 Then hintW = 60

    If hintW >= 218 Then
        m_LblHint1.Caption = "提示：双击列表项或按"
        m_LblHint1.Move pad, barTop + 5.5, 96, 13
        keyX = pad + 96
        m_BgKeyEnter.Move keyX, barTop + 3.5, 34, 16
        m_LblKeyEnter.Move keyX + 1, barTop + 5, 32, 12
        keyX = keyX + 37
        m_LblHint2.Move keyX, barTop + 5.5, 27, 13
        keyX = keyX + 27
        m_BgKeyEsc.Move keyX, barTop + 3.5, 26, 16
        m_LblKeyEsc.Move keyX + 1, barTop + 5, 24, 12
        keyX = keyX + 29
        m_LblHint3.Move keyX, barTop + 5.5, 24, 13

        m_BgKeyEnter.Visible = True
        m_LblKeyEnter.Visible = True
        m_LblHint2.Visible = True
        m_BgKeyEsc.Visible = True
        m_LblKeyEsc.Visible = True
        m_LblHint3.Visible = True
    Else
        m_LblHint1.Caption = "双击运行，Esc 关闭"
        m_LblHint1.Move pad, barTop + 5.5, hintW, 13
        m_BgKeyEnter.Visible = False
        m_LblKeyEnter.Visible = False
        m_LblHint2.Visible = False
        m_BgKeyEsc.Visible = False
        m_LblKeyEsc.Visible = False
        m_LblHint3.Visible = False
    End If

    ' 确保层级顺序
    m_BgSearch.ZOrder 1
    m_LblSearchIcon.ZOrder 0
    Me.txtSearch.ZOrder 0
    m_LblPlaceholder.ZOrder 0
    m_FlatClear.MoveTo searchLeft + searchW - 22, pad + 3, 18, 18

    m_BgListHeader.ZOrder 1
    m_LblListHeader.ZOrder 0
    m_LblListCount.ZOrder 0

    m_BgDetailCard.ZOrder 1
    m_BgCategory.ZOrder 0
    m_LblCategory.ZOrder 0
    m_LblDetailTitle.ZOrder 0
    m_DivDetail.ZOrder 0
    m_LblMacroCaption.ZOrder 0
    m_BgMacroBadge.ZOrder 0
    m_LblMacroBadge.ZOrder 0
    m_LblDescCaption.ZOrder 0
    Me.lblDetail.ZOrder 0

    m_BgKeyEnter.ZOrder 1
    m_LblKeyEnter.ZOrder 0
    m_BgKeyEsc.ZOrder 1
    m_LblKeyEsc.ZOrder 0

    UpdateScrollAreaAndRowWidths
    UpdateDetailBadgeSize innerW
    ResizeColumnWidths
    On Error GoTo 0
End Sub

Private Sub UpdateScrollAreaAndRowWidths()
    Dim countRows As Long
    Dim neededH As Single
    Dim needsScroll As Boolean
    Dim rowW As Single
    Dim rowItem As clsMacroListRow

    On Error Resume Next
    If m_MacroRows Is Nothing Then
        countRows = 0
    Else
        countRows = m_MacroRows.Count
    End If

    If countRows > 0 Then
        neededH = 4 + countRows * MACRO_ROW_STEP + 2
    Else
        neededH = 0
    End If

    needsScroll = (neededH > m_FrmList.InsideHeight)
    If needsScroll Then
        m_FrmList.ScrollHeight = neededH
        m_FrmList.ScrollBars = fmScrollBarsVertical
        rowW = m_FrmList.InsideWidth - 20
    Else
        m_FrmList.ScrollHeight = 0
        m_FrmList.ScrollBars = fmScrollBarsNone
        m_FrmList.ScrollTop = 0
        rowW = m_FrmList.InsideWidth - 8
    End If

    If rowW < 40 Then rowW = 40
    If Not m_MacroRows Is Nothing Then
        For Each rowItem In m_MacroRows
            rowItem.ResizeWidth rowW
        Next rowItem
    End If
    On Error GoTo 0
End Sub

Private Sub UpdateDetailBadgeSize(ByVal maxWidth As Single)
    Dim badgeW As Single
    Dim catW As Single
    Dim catText As String

    On Error Resume Next
    catText = m_LblCategory.Caption
    catW = Len(catText) * 9.5 + 18
    If catW < 46 Then catW = 46
    If catW > maxWidth Then catW = maxWidth
    m_BgCategory.Width = catW
    m_LblCategory.Width = catW

    If Len(m_CurrentEngName) = 0 Then
        badgeW = 48
    Else
        badgeW = Len(m_CurrentEngName) * 5.2 + 16
    End If
    If badgeW < 64 Then badgeW = 64
    If badgeW > maxWidth Then badgeW = maxWidth
    m_BgMacroBadge.Width = badgeW
    m_LblMacroBadge.Width = badgeW - 10
    On Error GoTo 0
End Sub

Private Sub ResizeColumnWidths()
    Dim w2 As Double

    On Error Resume Next
    w2 = Me.lstMacros.Width - 20
    If w2 < 0 Then w2 = 0
    Me.lstMacros.ColumnWidths = "0;" & w2 & ";0"
    On Error GoTo 0
End Sub

Private Sub SetDetailBody(ByVal bodyText As String)
    On Error Resume Next
    Me.lblDetail.Value = bodyText
    On Error GoTo 0
End Sub

Private Sub UpdatePlaceholderVisibility()
    On Error Resume Next
    m_LblPlaceholder.Visible = (Len(Me.txtSearch.Text) = 0 And Not m_SearchFocused)
    If Not m_FlatClear Is Nothing Then
        m_FlatClear.SetVisible (Len(Me.txtSearch.Text) > 0)
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' 交互与筛选逻辑
' ==============================================================================
Private Sub m_LblSearchIcon_Click()
    On Error Resume Next
    Me.txtSearch.SetFocus
    On Error GoTo 0
End Sub

Private Sub m_LblPlaceholder_Click()
    On Error Resume Next
    Me.txtSearch.SetFocus
    On Error GoTo 0
End Sub

Private Sub txtSearch_Enter()
    m_SearchFocused = True
    StyleFlatInputBackdrop m_BgSearch, True
    UpdatePlaceholderVisibility
End Sub

Private Sub txtSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    m_SearchFocused = False
    StyleFlatInputBackdrop m_BgSearch, False
    UpdatePlaceholderVisibility
End Sub

Private Sub txtSearch_Change()
    UpdatePlaceholderVisibility
    RefreshList Trim$(Me.txtSearch.Text)
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Me.lstMacros.ListCount > 0 Then
            If Me.lstMacros.ListIndex < Me.lstMacros.ListCount - 1 Then
                SelectRowIndex Me.lstMacros.ListIndex + 1
            End If
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        If Me.lstMacros.ListIndex > 0 Then
            SelectRowIndex Me.lstMacros.ListIndex - 1
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0
        btnRun_Click
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        If Len(Me.txtSearch.Text) > 0 Then
            Me.txtSearch.Text = vbNullString
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub lstMacros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        btnRun_Click
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub lstMacros_Change()
    If m_SyncingSelection Then Exit Sub
    UpdateDetailCard Me.lstMacros.ListIndex
End Sub

Private Sub lstMacros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnRun_Click
End Sub

Private Sub m_BtnClearSearch_Click()
    Me.txtSearch.Text = vbNullString
    On Error Resume Next
    Me.txtSearch.SetFocus
    On Error GoTo 0
End Sub

Private Sub m_BtnRefresh_Click()
    LoadDataToMemory
    Me.txtSearch.Text = vbNullString
End Sub

Private Sub LoadDataToMemory()
    On Error Resume Next
    m_AllMacros = Application.Run("GetMyMacroRegistry")
    On Error GoTo 0
    RefreshList vbNullString
End Sub

Private Sub ClearMacroRows()
    Dim rowItem As clsMacroListRow
    Dim i As Long

    On Error Resume Next
    If Not m_MacroRows Is Nothing Then
        For Each rowItem In m_MacroRows
            rowItem.Dispose
        Next rowItem
        Set m_MacroRows = New Collection
    End If
    If Not m_FrmList Is Nothing Then
        For i = m_FrmList.Controls.Count - 1 To 0 Step -1
            m_FrmList.Controls.Remove m_FrmList.Controls(i).Name
        Next i
    End If
    On Error GoTo 0
End Sub

Private Sub RefreshList(ByVal keyword As String)
    Dim i As Long
    Dim totalCount As Long
    Dim bMatch As Boolean
    Dim sEng As String
    Dim sCn As String
    Dim sDesc As String
    Dim rowIdx As Long
    Dim rowItem As clsMacroListRow
    Dim rowW As Single

    Me.lstMacros.Clear
    ClearMacroRows

    If IsEmpty(m_AllMacros) Then
        m_LblListCount.Caption = "0 项"
        UpdateDetailCard -1
        UpdateScrollAreaAndRowWidths
        Exit Sub
    End If

    totalCount = UBound(m_AllMacros) - LBound(m_AllMacros) + 1
    rowW = m_FrmList.InsideWidth - 20
    If rowW < 40 Then rowW = 40

    For i = LBound(m_AllMacros) To UBound(m_AllMacros)
        sEng = m_AllMacros(i)(0)
        sCn = m_AllMacros(i)(1)
        sDesc = m_AllMacros(i)(2)

        If keyword = vbNullString Then
            bMatch = True
        ElseIf InStr(1, sEng, keyword, vbTextCompare) > 0 Or _
               InStr(1, sCn, keyword, vbTextCompare) > 0 Or _
               InStr(1, sDesc, keyword, vbTextCompare) > 0 Then
            bMatch = True
        Else
            bMatch = False
        End If

        If bMatch Then
            rowIdx = Me.lstMacros.ListCount
            Me.lstMacros.AddItem
            Me.lstMacros.List(rowIdx, 0) = sEng
            Me.lstMacros.List(rowIdx, 1) = sCn
            Me.lstMacros.List(rowIdx, 2) = sDesc

            Set rowItem = New clsMacroListRow
            rowItem.Init m_FrmList, rowIdx, sCn, 4 + rowIdx * MACRO_ROW_STEP, rowW, Me
            m_MacroRows.Add rowItem
        End If
    Next i

    If keyword = vbNullString Then
        m_LblListCount.Caption = "全部 " & Me.lstMacros.ListCount & " 项"
    Else
        m_LblListCount.Caption = Me.lstMacros.ListCount & " / " & totalCount & " 项"
    End If

    UpdateScrollAreaAndRowWidths

    If Me.lstMacros.ListCount > 0 Then
        SelectRowIndex 0
    Else
        UpdateDetailCard -1
    End If
End Sub

Private Sub SelectRowIndex(ByVal idx As Long)
    Dim rowItem As clsMacroListRow
    Dim rowTop As Single
    Dim rowBottom As Single

    If idx < 0 Or idx >= Me.lstMacros.ListCount Then
        UpdateDetailCard -1
        Exit Sub
    End If

    m_SyncingSelection = True
    On Error Resume Next
    Me.lstMacros.Selected(idx) = True
    On Error GoTo 0
    m_SyncingSelection = False

    If Not m_MacroRows Is Nothing Then
        For Each rowItem In m_MacroRows
            rowItem.SetSelected (rowItem.RowIndex = idx)
        Next rowItem
    End If

    On Error Resume Next
    If m_FrmList.ScrollBars = fmScrollBarsVertical Then
        rowTop = 4 + idx * MACRO_ROW_STEP
        rowBottom = rowTop + MACRO_ROW_HEIGHT
        If rowTop < m_FrmList.ScrollTop Then
            m_FrmList.ScrollTop = rowTop - 4
        ElseIf rowBottom > m_FrmList.ScrollTop + m_FrmList.InsideHeight Then
            m_FrmList.ScrollTop = rowBottom - m_FrmList.InsideHeight + 4
        End If
    End If
    On Error GoTo 0

    UpdateDetailCard idx
End Sub

Private Sub UpdateDetailCard(ByVal idx As Long)
    Dim engName As String
    Dim cnTitle As String
    Dim cnDesc As String
    Dim maxW As Single

    maxW = m_BgDetailCard.Width - 28
    If maxW < 80 Then maxW = 80

    If idx >= 0 And idx < Me.lstMacros.ListCount Then
        engName = Me.lstMacros.List(idx, 0)
        cnTitle = Me.lstMacros.List(idx, 1)
        cnDesc = Me.lstMacros.List(idx, 2)
        m_CurrentEngName = engName
        m_LblCategory.Caption = GetMacroCategory(engName)
        m_LblDetailTitle.Caption = cnTitle
        m_LblMacroBadge.Caption = engName
        SetDetailBody cnDesc
    Else
        m_CurrentEngName = vbNullString
        m_LblCategory.Caption = "无匹配"
        m_LblDetailTitle.Caption = "未选中任何功能"
        m_LblMacroBadge.Caption = "-"
        SetDetailBody "请从左侧列表选择功能，或调整上方搜索关键词。"
    End If

    UpdateDetailBadgeSize maxW
End Sub

Private Function GetMacroCategory(ByVal engName As String) As String
    Select Case engName
        Case "SetHyperlinksAndFieldsToBlue", "Wrapper_PasteRelinkCrossReference", "Wrapper_GenerateCaptionReferencePair"
            GetMacroCategory = "域与格式"
        Case "BatchConvertWordToPDF"
            GetMacroCategory = "批量转换"
        Case "BatchRenameFiles"
            GetMacroCategory = "文件批处理"
        Case "RenameCurrentDocument"
            GetMacroCategory = "文件操作"
        Case "ConvertHeadingNumbers", "RemoveUnusedStyles"
            GetMacroCategory = "样式排版"
        Case "BatchSetMargins", "LinkToThePreviousSection"
            GetMacroCategory = "页面设置"
        Case "BatchAutoFitTablesToWindow"
            GetMacroCategory = "表格排版"
        Case "BatchAcceptAndClean", "Wrapper_AcceptFormatRevisions", "Wrapper_ShowReviewerFilter"
            GetMacroCategory = "审阅清理"
        Case "BatchDetectCleanVersions", "BatchDetectHighlights"
            GetMacroCategory = "质量检查"
        Case "ExtractAbbreviations", "NormalizeScientificTerms"
            GetMacroCategory = "文本工具"
        Case Else
            GetMacroCategory = "实用宏"
    End Select
End Function

Private Sub btnRun_Click()
    If Me.lstMacros.ListIndex = -1 Then
        MsgBox "请先选择一个功能！", vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrH
    Application.Run Me.lstMacros.List(Me.lstMacros.ListIndex, 0)
    Exit Sub
ErrH:
    MsgBox "运行出错：" & Err.Description, vbCritical
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ClearMacroRows
    ClearHoverButtons m_HoverButtons
    Set m_FlatClear = Nothing
    Set m_FlatRefresh = Nothing
    Set m_FlatClose = Nothing
    Set m_FlatRun = Nothing
    Set m_LblSearchIcon = Nothing
    Set m_LblPlaceholder = Nothing
    Set m_FrmList = Nothing
End Sub
