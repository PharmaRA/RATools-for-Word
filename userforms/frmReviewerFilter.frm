VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReviewerFilter 
   Caption         =   "按审阅者筛选"
   ClientHeight    =   5832
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6180
   OleObjectBlob   =   "frmReviewerFilter.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmReviewerFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================
' 按审阅者筛选面板
' 勾选即生效，底层与 Word 原生「显示标记 -> 特定人员」等效。
'
' 界面是一组真正的复选框（按作者数在运行时生成）：
'   - 第一项「全部审阅者」：全选 / 全不选；
'   - 其余每项一位作者，勾选状态就是文档里的显示状态；
'   - 单击某位作者后，「只看此人」按钮可一键只显示他。
' 动态复选框的 Click 由 clsReviewerItem 接收后转回本窗体，
' 勾选状态只有 m_Checked 一个来源，「全部审阅者」每次都由
' 各行状态重新推导，不会再出现上下不一致。
' =============================================

Private Const ROW_HEIGHT As Single = 26
Private Const ROW_TOP As Single = 8
Private Const MAX_FRAME_HEIGHT As Single = 300

Private m_Authors() As String
Private m_Counts() As Long
Private m_Checked() As Boolean
Private m_AuthorCount As Long
Private m_DocKey As String
Private m_Updating As Boolean
Private m_Items As Collection
Private m_LastClickedIndex As Long

Private Sub UserForm_Initialize()
    m_DocKey = ""
    m_LastClickedIndex = -1

    Me.caption = REVIEWER_FILTER_CAPTION
    Me.btnClose.caption = "关闭"
    Me.btnOnlyThis.caption = "只看此人"

    With Me.lblHint.Font
        .Name = "Microsoft YaHei"
        .Size = 9
    End With
    Me.lblHint.ForeColor = RGB(110, 110, 110)

    With Me.btnClose.Font
        .Name = "Microsoft YaHei"
        .Size = 9
    End With
    With Me.btnOnlyThis.Font
        .Name = "Microsoft YaHei"
        .Size = 9
    End With

    With Me.frmList
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(210, 214, 220)
        .ScrollBars = fmScrollBarsNone
    End With

    RefreshReviewerList False
End Sub

' 打开面板时把焦点放到关闭按钮，避免第一个复选框带虚线焦点框
Private Sub UserForm_Activate()
    On Error Resume Next
    Me.btnClose.SetFocus
    On Error GoTo 0
End Sub

' 当前面板对应的文档键（供模块判断是否需要重建列表）
Public Property Get FilteredDocumentKey() As String
    FilteredDocumentKey = m_DocKey
End Property

' 读取审阅者并重建面板。forceRefresh = True 时重新枚举文档中的作者；
' 否则优先复用模块级缓存，窗口激活等高频路径不再重复扫描修订。
Public Sub RefreshReviewerList(Optional ByVal forceRefresh As Boolean = False)
    Dim i As Long
    Dim activeKey As String

    If Application.Documents.Count = 0 Then
        m_AuthorCount = 0
        m_DocKey = ""
        ClearRows
        Me.lblHint.caption = "当前没有打开的文档。"
        Exit Sub
    End If

    activeKey = GetDocumentCacheKey(ActiveDocument)

    ' 面板已是当前文档的内容：不重建（这是"每次切回 Word 都转圈"的根因）
    If Not forceRefresh Then
        If StrComp(activeKey, m_DocKey, vbBinaryCompare) = 0 And Not m_Items Is Nothing Then
            Exit Sub
        End If
    End If

    m_DocKey = activeKey
    m_LastClickedIndex = -1

    GetReviewerSnapshot ActiveDocument, m_Authors, m_Counts, m_AuthorCount, forceRefresh

    If m_AuthorCount > 0 Then
        ReDim m_Checked(0 To m_AuthorCount - 1)
    Else
        ReDim m_Checked(0 To 0)
    End If

    For i = 0 To m_AuthorCount - 1
        m_Checked(i) = IsReviewerVisible(ActiveDocument, m_Authors(i))
    Next i

    BuildRows
    UpdateHint
End Sub

' 按当前作者数重建复选框
Private Sub BuildRows()
    Dim i As Long
    Dim innerWidth As Single
    Dim chk As MSForms.CheckBox

    m_Updating = True

    ClearRows
    Set m_Items = New Collection

    If m_AuthorCount > 0 Then
        innerWidth = Me.frmList.InsideWidth - 24
        If innerWidth < 120 Then innerWidth = 120

        ' 第一项：全部审阅者
        Set chk = AddRow(-1, "全部审阅者（" & m_AuthorCount & " 位）", innerWidth)
        chk.Font.Bold = True

        ' 其余：每位作者
        For i = 0 To m_AuthorCount - 1
            Set chk = AddRow(i, m_Authors(i) & "（" & m_Counts(i) & " 处）", innerWidth)
        Next i

        RefreshChecksFromFlags
    End If

    LayoutRows
    UpdateOnlyButton
    m_Updating = False
End Sub

Private Function AddRow(ByVal authorIndex As Long, ByVal caption As String, _
                        ByVal innerWidth As Single) As MSForms.CheckBox
    Dim chk As MSForms.CheckBox
    Dim item As clsReviewerItem
    Dim row As Long

    ' 行号：全部审阅者占第 0 行
    If authorIndex < 0 Then
        row = 0
    Else
        row = authorIndex + 1
    End If

    Set chk = Me.frmList.Controls.Add("Forms.CheckBox.1", "chkRow" & row, True)
    With chk
        .caption = caption
        .Left = 14
        .Top = ROW_TOP + row * ROW_HEIGHT
        .Width = innerWidth
        .Height = ROW_HEIGHT
        .Font.Name = "Microsoft YaHei"
        .Font.Size = 9.5
        .ForeColor = RGB(32, 32, 32)
        .BackColor = RGB(255, 255, 255)
        .BackStyle = fmBackStyleOpaque
        .Value = False
    End With

    Set item = New clsReviewerItem
    Set item.chk = chk
    Set item.Owner = Me
    item.authorIndex = authorIndex
    m_Items.Add item

    Set AddRow = chk
End Function

' 按实际行数调整列表区与窗体高度：人少时不留大片空白，人多时滚动。
' 内容高度取最后一行控件的实际下沿——不同 DPI 下复选框实际高度会大于 ROW_HEIGHT。
Private Sub LayoutRows()
    Dim needed As Single
    Dim lastControl As Control
    Dim frameHeight As Single
    Dim chrome As Single

    If m_AuthorCount <= 0 Then Exit Sub

    Set lastControl = Me.frmList.Controls(Me.frmList.Controls.Count - 1)
    needed = lastControl.Top + lastControl.Height + 8

    frameHeight = needed
    If frameHeight > MAX_FRAME_HEIGHT Then frameHeight = MAX_FRAME_HEIGHT

    ' 改窗体高度前先量出标题栏 + 边框的厚度
    chrome = Me.Height - Me.InsideHeight

    Me.frmList.Height = frameHeight

    ' 边框会吃掉几像素：先把外层高度补足，让最后一行完整露出
    If needed > Me.frmList.InsideHeight Then
        frameHeight = frameHeight + (needed - Me.frmList.InsideHeight) + 2
        If frameHeight > MAX_FRAME_HEIGHT Then frameHeight = MAX_FRAME_HEIGHT
        Me.frmList.Height = frameHeight
    End If

    ' 内容装得下就不要滚动条（ScrollBars 不会自动隐藏），装不下才启用
    If needed > Me.frmList.InsideHeight Then
        Me.frmList.ScrollHeight = needed
        Me.frmList.ScrollBars = fmScrollBarsVertical
    Else
        Me.frmList.ScrollHeight = 0
        Me.frmList.ScrollBars = fmScrollBarsNone
    End If

    Me.btnOnlyThis.Top = Me.frmList.Top + Me.frmList.Height + 12
    Me.btnClose.Top = Me.btnOnlyThis.Top
    Me.Height = Me.btnOnlyThis.Top + Me.btnOnlyThis.Height + 12 + chrome
End Sub

Private Sub ClearRows()
    On Error Resume Next
    If Not m_Items Is Nothing Then
        Do While m_Items.Count > 0
            Me.frmList.Controls.Remove m_Items(1).chk.Name
            m_Items.Remove 1
        Loop
        Set m_Items = Nothing
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------
' 由 clsReviewerItem 回调（动态控件的 Click 无法直接在窗体上接收）
' ---------------------------------------------------------------

Public Sub OnItemClicked(ByVal authorIndex As Long, ByVal newValue As Boolean)
    If m_Updating Then Exit Sub

    If authorIndex < 0 Then
        HandleAllClicked newValue
    Else
        HandleAuthorClicked authorIndex, newValue
    End If
End Sub

' 「全部审阅者」复选框被点击
Private Sub HandleAllClicked(ByVal newValue As Boolean)
    Dim i As Long
    Dim target As Boolean

    If m_AuthorCount = 0 Then Exit Sub

    m_LastClickedIndex = -1

    target = newValue
    If target = AllFlagsSet() Then
        ' 已经是目标状态：可能只是控件自己先翻了一下，这里按状态回填以避免抖动
        RefreshChecksFromFlags
        UpdateOnlyButton
        Exit Sub
    End If

    For i = 0 To m_AuthorCount - 1
        m_Checked(i) = target
    Next i

    If Not ApplyAllAuthors() Then
        Me.lblHint.caption = "无法应用筛选：请确认文档处于页面视图等可编辑视图，且未启用阅读模式。"
        RefreshChecksFromFlags
        UpdateOnlyButton
        Exit Sub
    End If

    RefreshChecksFromFlags
    UpdateHint
    UpdateOnlyButton
End Sub

' 某位作者的复选框被点击
Private Sub HandleAuthorClicked(ByVal idx As Long, ByVal newValue As Boolean)
    If idx < 0 Or idx >= m_AuthorCount Then Exit Sub
    If m_Checked(idx) = newValue Then Exit Sub

    If Not SetAuthorVisible(idx, newValue) Then
        Me.lblHint.caption = "无法应用筛选：请确认文档处于页面视图等可编辑视图，且未启用阅读模式。"
        RefreshChecksFromFlags
        UpdateOnlyButton
        Exit Sub
    End If

    m_Checked(idx) = newValue
    m_LastClickedIndex = idx
    RefreshChecksFromFlags
    UpdateHint
    UpdateOnlyButton
End Sub

' ---------------------------------------------------------------
' 对外 API（自动化测试与外部调用）
' ---------------------------------------------------------------

' 切换某位审阅者的显示状态
Public Sub ToggleAuthorAt(ByVal idx As Long)
    Dim newValue As Boolean

    If m_Updating Then Exit Sub
    If idx < 0 Or idx >= m_AuthorCount Then Exit Sub

    newValue = Not m_Checked(idx)

    If Not SetAuthorVisible(idx, newValue) Then
        Me.lblHint.caption = "无法应用筛选：请确认文档处于页面视图等可编辑视图，且未启用阅读模式。"
        Exit Sub
    End If

    m_Checked(idx) = newValue
    m_LastClickedIndex = idx
    RefreshChecksFromFlags
    UpdateHint
    UpdateOnlyButton
End Sub

' 只看某位审阅者
Public Sub ShowOnlyAuthorAt(ByVal idx As Long)
    Dim i As Long

    If m_Updating Then Exit Sub
    If idx < 0 Or idx >= m_AuthorCount Then Exit Sub

    For i = 0 To m_AuthorCount - 1
        If i = idx Then
            m_Checked(i) = True
        Else
            m_Checked(i) = False
        End If
    Next i

    If Not ApplyAllAuthors() Then
        Me.lblHint.caption = "无法应用筛选：请确认文档处于页面视图等可编辑视图，且未启用阅读模式。"
        Exit Sub
    End If

    m_LastClickedIndex = idx
    RefreshChecksFromFlags
    UpdateHint
    UpdateOnlyButton
End Sub

' 全部显示 / 全部隐藏
Public Sub ToggleAllReviewers()
    HandleAllClicked Not AllFlagsSet()
End Sub

' 只看最近单击的那位审阅者（「只看此人」按钮）
Public Sub ShowOnlyLastClicked()
    If m_AuthorCount = 0 Then Exit Sub

    If m_LastClickedIndex < 0 Or m_LastClickedIndex >= m_AuthorCount Then
        Me.lblHint.caption = "请先单击一位审阅者，再点「只看此人」。"
        Exit Sub
    End If

    ShowOnlyAuthorAt m_LastClickedIndex
End Sub

Private Sub btnOnlyThis_Click()
    ShowOnlyLastClicked
End Sub

' ---------------------------------------------------------------
' 状态回填
' ---------------------------------------------------------------

' 按 m_Checked 回填所有复选框（第 1 项是「全部审阅者」，单独算）
Private Sub RefreshChecksFromFlags()
    Dim i As Long

    If m_Items Is Nothing Then Exit Sub

    For i = 2 To m_Items.Count
        m_Items(i).SyncValue m_Checked(m_Items(i).authorIndex)
    Next i
    If m_AuthorCount > 0 Then m_Items(1).SyncValue AllFlagsSet()
End Sub

Private Sub UpdateHint()
    Dim visibleCount As Long
    Dim i As Long

    If m_AuthorCount <= 0 Then
        Me.lblHint.caption = "本文档中没有批注或修订。"
        Exit Sub
    End If

    For i = 0 To m_AuthorCount - 1
        If m_Checked(i) Then visibleCount = visibleCount + 1
    Next i

    Me.lblHint.caption = "显示 " & visibleCount & "/" & m_AuthorCount & " 位审阅者"
End Sub

' 「只看此人」需要先选中一位作者
Private Sub UpdateOnlyButton()
    On Error Resume Next
    Me.btnOnlyThis.Enabled = (m_AuthorCount > 0 And m_LastClickedIndex >= 0 And m_LastClickedIndex < m_AuthorCount)
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' 供外部（自动化测试）关闭面板
Public Sub ClosePanel()
    Unload Me
End Sub

' ---------------------------------------------------------------
' 文档写入
' ---------------------------------------------------------------

' 把某一位作者的显示状态写入文档（文档已切换时先刷新面板并放弃本次操作）
Private Function SetAuthorVisible(ByVal idx As Long, ByVal newState As Boolean) As Boolean
    If Not EnsureSameDocument() Then Exit Function
    If idx < 0 Or idx >= m_AuthorCount Then Exit Function
    SetAuthorVisible = SetReviewerVisible(ActiveDocument, m_Authors(idx), newState)
End Function

Private Function ApplyAllAuthors() As Boolean
    If Not EnsureSameDocument() Then Exit Function
    ApplyAllAuthors = ApplyReviewerFilter(ActiveDocument, m_Authors, m_Checked, m_AuthorCount)
End Function

Private Function EnsureSameDocument() As Boolean
    If Application.Documents.Count = 0 Then Exit Function
    If StrComp(GetDocumentCacheKey(ActiveDocument), m_DocKey, vbBinaryCompare) <> 0 Then
        RefreshReviewerList False
        Exit Function
    End If
    EnsureSameDocument = True
End Function

Private Function AllFlagsSet() As Boolean
    Dim i As Long

    If m_AuthorCount <= 0 Then Exit Function
    For i = 0 To m_AuthorCount - 1
        If Not m_Checked(i) Then Exit Function
    Next i
    AllFlagsSet = True
End Function


