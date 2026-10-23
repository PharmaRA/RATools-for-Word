VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReviewerFilter 
   Caption         =   "按审阅者筛选"
   ClientHeight    =   6432
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6588
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
' 交互约定：
'   - 列表为复选框列表，可随意多选；
'   - 「全部审阅者」一次全选 / 全不选，不承载其他含义；
'   - 双击某位审阅者 = 只看该审阅者。
' 事件过程只做转发，实际逻辑放在 Public 过程里，便于自动化测试直接调用。
' =============================================

Private m_Authors() As String
Private m_Counts() As Long
Private m_Checked() As Boolean
Private m_AuthorCount As Long
Private m_DocKey As String
Private m_Updating As Boolean

Private Sub UserForm_Initialize()
    m_DocKey = ""

    Me.Caption = REVIEWER_FILTER_CAPTION
    Me.chkAll.Caption = "全部审阅者"
    Me.btnClose.Caption = "关闭"

    Me.lstAuthors.ListStyle = fmListStyleOption
    Me.lstAuthors.MultiSelect = fmMultiSelectMulti

    RefreshReviewerList
End Sub

' 按当前文档重新读取审阅者列表与显示状态
Public Sub RefreshReviewerList()
    Dim i As Long
    Dim visibleCount As Long
    Dim doc As Document

    If Application.Documents.Count = 0 Then
        Me.lstAuthors.Clear
        Me.chkAll.Enabled = False
        Me.lblHint.Caption = "当前没有打开的文档。"
        m_AuthorCount = 0
        Exit Sub
    End If

    Set doc = ActiveDocument
    m_DocKey = doc.FullName

    m_Updating = True

    m_Authors = CollectReviewerNames(doc, m_Counts, m_AuthorCount)
    If m_AuthorCount > 0 Then
        ReDim m_Checked(0 To m_AuthorCount - 1)
    Else
        ReDim m_Checked(0 To 0)
    End If

    Me.lstAuthors.Clear
    For i = 0 To m_AuthorCount - 1
        Me.lstAuthors.AddItem BuildReviewerLabel(m_Authors(i), m_Counts(i))
        m_Checked(i) = IsReviewerVisible(doc, m_Authors(i))
        Me.lstAuthors.Selected(i) = m_Checked(i)
        If m_Checked(i) Then visibleCount = visibleCount + 1
    Next i

    Me.chkAll.Enabled = (m_AuthorCount > 0)
    Me.chkAll.Value = (m_AuthorCount > 0 And visibleCount = m_AuthorCount)
    Me.lblHint.Caption = GetListSummary()

    m_Updating = False
End Sub

Private Function BuildReviewerLabel(ByVal author As String, ByVal markCount As Long) As String
    BuildReviewerLabel = author & "（" & markCount & " 处）"
End Function

' 单击审阅者：切换其显示状态
Public Sub ToggleAuthorAt(ByVal idx As Long)
    If m_Updating Or m_AuthorCount = 0 Then Exit Sub
    If idx < 0 Or idx >= m_AuthorCount Then Exit Sub

    ' 不同 MSForms 版本对"点击文字"是否自动切换勾选的行为不一致；
    ' 若点击后状态没有变化，由代码补一次切换。
    If Me.lstAuthors.Selected(idx) = m_Checked(idx) Then
        m_Updating = True
        Me.lstAuthors.Selected(idx) = Not m_Checked(idx)
        m_Updating = False
    End If

    SyncCheckedFromList
    Me.chkAll.Value = AllChecked()
    ApplyCurrentFilter
End Sub

' 双击审阅者：只看该审阅者
Public Sub ShowOnlyAuthorAt(ByVal idx As Long)
    Dim i As Long

    If m_Updating Or m_AuthorCount = 0 Then Exit Sub
    If idx < 0 Or idx >= m_AuthorCount Then Exit Sub

    For i = 0 To m_AuthorCount - 1
        If i = idx Then
            m_Checked(i) = True
        Else
            m_Checked(i) = False
        End If
    Next i

    ApplyCheckedFlags
End Sub

' 「全部审阅者」：一次全选或全不选
Public Sub ToggleAllReviewers()
    Dim i As Long
    Dim target As Boolean

    If m_Updating Or m_AuthorCount = 0 Then Exit Sub

    target = (Me.chkAll.Value = True)

    ' 程序化回填 chkAll 也会触发 Click；目标状态与当前勾选一致时说明
    ' 这次不是用户意图，直接返回，避免覆盖刚才算好的筛选结果。
    If target = AllFlagsSet() Then Exit Sub

    For i = 0 To m_AuthorCount - 1
        m_Checked(i) = target
    Next i

    ApplyCheckedFlags
End Sub

' 按 m_Checked 写回显示状态，并重建列表使其勾选状态与之一致。
' MSForms 复选框列表在窗体未显示时逐个赋值 Selected 并不可靠，
' 因此程序化改状态后统一用 Clear + AddItem 重建（与 RefreshReviewerList 同路径）。
Private Sub ApplyCheckedFlags()
    If Not ApplyFilterToDocument() Then Exit Sub

    SetAllCheckboxState AllFlagsSet()
    RebuildListFromFlags
End Sub

' 代码里给 CheckBox.Value 赋值会触发 Click（MSForms 的既有行为，且事件可能
' 在本过程返回后才派发，单靠 m_Updating 挡不住）。ToggleAllReviewers 用
' "目标状态是否已等于当前勾选"来识别这类回填事件，所以这里只需赋值。
Private Sub SetAllCheckboxState(ByVal state As Boolean)
    m_Updating = True
    Me.chkAll.Value = state
    m_Updating = False
End Sub

Private Sub RebuildListFromFlags()
    Dim i As Long

    m_Updating = True
    Me.lstAuthors.Clear
    For i = 0 To m_AuthorCount - 1
        Me.lstAuthors.AddItem BuildReviewerLabel(m_Authors(i), m_Counts(i))
        Me.lstAuthors.Selected(i) = m_Checked(i)
    Next i
    m_Updating = False

    Me.lblHint.Caption = GetListSummary()
End Sub

Private Sub lstAuthors_Click()
    ToggleAuthorAt Me.lstAuthors.ListIndex
End Sub

Private Sub lstAuthors_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ShowOnlyAuthorAt Me.lstAuthors.ListIndex
End Sub

' 用 MouseUp 而不是 Click：程序里给 .Value 赋值会触发 Click（且事件可能延迟到
' 下一次消息循环才派发），会把刚算好的勾选覆盖成"全选/全不选"。鼠标抬起只有
' 真实点击才会有，天然区分用户操作与程序化回填。
Private Sub chkAll_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                           ByVal X As Single, ByVal Y As Single)
    ToggleAllReviewers
End Sub

' 键盘操作（Tab 到勾选框后按空格）不会触发 MouseUp，需要单独接一次
Private Sub chkAll_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeySpace Then ToggleAllReviewers
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' 供外部（Ribbon 再次点击、自动化测试）关闭面板
Public Sub ClosePanel()
    Unload Me
End Sub

Private Sub SyncCheckedFromList()
    Dim i As Long

    If m_AuthorCount > 0 Then
        ReDim m_Checked(0 To m_AuthorCount - 1)
    Else
        ReDim m_Checked(0 To 0)
    End If

    For i = 0 To m_AuthorCount - 1
        m_Checked(i) = Me.lstAuthors.Selected(i)
    Next i
End Sub

Private Function AllChecked() As Boolean
    Dim i As Long

    If m_AuthorCount <= 0 Then Exit Function
    For i = 0 To m_AuthorCount - 1
        If Not Me.lstAuthors.Selected(i) Then Exit Function
    Next i
    AllChecked = True
End Function

' 勾选标志是否全部为真（用于回填「全部审阅者」）
Private Function AllFlagsSet() As Boolean
    Dim i As Long

    If m_AuthorCount <= 0 Then Exit Function
    For i = 0 To m_AuthorCount - 1
        If Not m_Checked(i) Then Exit Function
    Next i
    AllFlagsSet = True
End Function

' 把 m_Checked 应用到文档；文档已切换时先刷新列表并放弃本次应用
Private Function ApplyFilterToDocument() As Boolean
    If Application.Documents.Count = 0 Then Exit Function

    If ActiveDocument.FullName <> m_DocKey Then
        RefreshReviewerList
        Exit Function
    End If

    If Not ApplyReviewerFilter(ActiveDocument, m_Authors, m_Checked, m_AuthorCount) Then
        ' 勾选动作不弹模态框，失败信息直接写在面板上，避免打断连续勾选
        Me.lblHint.Caption = "无法应用筛选：请确认文档处于页面视图等可编辑视图，且未启用阅读模式。"
        Exit Function
    End If

    ApplyFilterToDocument = True
End Function

' 单击/双击/全选后统一走这里：以列表勾选为准，写回文档并同步汇总文案
Private Sub ApplyCurrentFilter()
    If Application.Documents.Count = 0 Then Exit Sub

    If ActiveDocument.FullName <> m_DocKey Then
        RefreshReviewerList
        Exit Sub
    End If

    If ApplyFilterToDocument() Then Me.lblHint.Caption = GetListSummary()
End Sub

Private Function GetListSummary() As String
    If m_AuthorCount <= 0 Then
        GetListSummary = "本文档中没有批注或修订。"
    Else
        GetListSummary = "共 " & m_AuthorCount & " 位审阅者，勾选立即生效；双击某人可只看该审阅者。"
    End If
End Function


