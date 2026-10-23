Attribute VB_Name = "Macro_ReviewerFilter"
Option Explicit

' =============================================
' 模块名：Macro_ReviewerFilter
' 功能：按审阅者筛选文档中显示的批注与修订。
' 实现：Word 对象模型没有"按作者过滤"的属性，但
'       ActiveWindow.View.Reviewers(姓名).Visible 与原生
'       「显示标记 -> 特定人员」菜单完全等效。本模块负责收集
'       作者、读写显示状态，界面在 frmReviewerFilter。
' =============================================

Private Const DIALOG_TITLE As String = "按审阅者筛选"

' 窗体 Caption，必须与 frmReviewerFilter 的 Me.Caption 一致
Public Const REVIEWER_FILTER_CAPTION As String = "按审阅者筛选"

' Ribbon 回调：打开审阅者筛选面板
Public Sub ShowReviewerFilter(control As IRibbonControl)
    Dim doc As Document
    Dim authorNames() As String
    Dim authorCounts() As Long
    Dim total As Long

    If Application.Documents.Count = 0 Then
        MsgBox "当前没有打开的文档。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    Set doc = ActiveDocument
    authorNames = CollectReviewerNames(doc, authorCounts, total)

    If total = 0 Then
        MsgBox "当前文档中没有批注或修订，无需筛选。", vbInformation, DIALOG_TITLE
        Exit Sub
    End If

    On Error GoTo ErrH

    ' 面板已经打开时只刷新内容，避免重复 Show
    If IsUserFormWindowVisible(REVIEWER_FILTER_CAPTION) Then
        frmReviewerFilter.RefreshReviewerList
        Exit Sub
    End If

    Load frmReviewerFilter
    frmReviewerFilter.Show vbModeless
    Exit Sub

ErrH:
    MsgBox "无法打开审阅者筛选面板：" & Err.Description, vbExclamation, DIALOG_TITLE
End Sub

' 宏列表入口：Application.Run 不支持带参数的过程，包装为无参入口
Public Sub Wrapper_ShowReviewerFilter()
    ShowReviewerFilter Nothing
End Sub

' 切换或激活文档后刷新面板内容（面板未显示时什么都不做）
Public Sub RefreshReviewerFilterIfVisible()
    On Error Resume Next
    If Not IsUserFormWindowVisible(REVIEWER_FILTER_CAPTION) Then Exit Sub
    frmReviewerFilter.RefreshReviewerList
End Sub

' =============================================
' 审阅者收集
' 作者名按 Word 的认定原样保留（大小写敏感），与原生菜单一致；
' 返回按姓名升序排列的数组，authorCounts 为对应标记处数。
' =============================================
Public Function CollectReviewerNames(ByVal doc As Document, _
                                    ByRef authorCounts() As Long, _
                                    ByRef total As Long) As String()
    Dim dict As Object
    Dim story As Range
    Dim cmt As Comment
    Dim authorNames() As String
    Dim i As Long

    On Error GoTo Failed
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 0 ' Binary：Word 把大小写不同的作者视为不同审阅者

    ' 正文及各页眉页脚、脚注尾注、文本框等故事区
    For Each story In doc.StoryRanges
        CollectRevisionsFromStory story, dict
    Next story

    ' 批注：Word 的 Comments 集合同时包含批注回复
    For Each cmt In doc.Comments
        AddAuthorCount dict, cmt.Author
    Next cmt

    total = dict.Count
    If total = 0 Then Exit Function

    authorNames = SortedAuthorNames(dict)
    ReDim authorCounts(0 To total - 1)
    For i = 0 To total - 1
        authorCounts(i) = CLng(dict(authorNames(i)))
    Next i

    CollectReviewerNames = authorNames
    Exit Function

Failed:
    total = 0
End Function

' 逐级遍历同一故事类型的 NextStoryRange（各节页眉页脚等）
Private Sub CollectRevisionsFromStory(ByVal storyRange As Range, ByVal dict As Object)
    Dim current As Range
    Dim rev As Revision

    Set current = storyRange
    Do While Not current Is Nothing
        For Each rev In current.Revisions
            AddAuthorCount dict, rev.Author
        Next rev
        Set current = current.NextStoryRange
    Loop
End Sub

Private Sub AddAuthorCount(ByVal dict As Object, ByVal author As String)
    If Len(author) = 0 Then Exit Sub
    If dict.Exists(author) Then
        dict(author) = CLng(dict(author)) + 1
    Else
        dict.Add author, 1
    End If
End Sub

' Scripting.Dictionary 不提供排序，这里用插入排序按姓名升序输出
Private Function SortedAuthorNames(ByVal dict As Object) As String()
    Dim result() As String
    Dim i As Long
    Dim j As Long
    Dim key As Variant
    Dim current As String

    ReDim result(0 To dict.Count - 1)
    i = 0
    For Each key In dict.Keys
        result(i) = CStr(key)
        i = i + 1
    Next key

    For i = 1 To UBound(result)
        current = result(i)
        j = i - 1
        Do While j >= 0
            If StrComp(result(j), current, vbTextCompare) <= 0 Then Exit Do
            result(j + 1) = result(j)
            j = j - 1
        Loop
        result(j + 1) = current
    Next i

    SortedAuthorNames = result
End Function

' =============================================
' 显示状态读写
' =============================================

' 某位审阅者当前是否显示；读取失败按显示处理
Public Function IsReviewerVisible(ByVal doc As Document, ByVal author As String) As Boolean
    On Error GoTo DefaultVisible
    IsReviewerVisible = doc.ActiveWindow.View.Reviewers.Item(author).Visible
    Exit Function

DefaultVisible:
    IsReviewerVisible = True
End Function

' 按勾选结果逐位写入显示状态。文档中已不存在的作者跳过，不影响其余作者。
Public Function ApplyReviewerFilter(ByVal doc As Document, _
                                    ByRef authorNames() As String, _
                                    ByRef visibleFlags() As Boolean, _
                                    ByVal total As Long) As Boolean
    Dim win As Window
    Dim i As Long

    On Error GoTo Failed
    Set win = doc.ActiveWindow
    For i = 0 To total - 1
        On Error Resume Next
        If Len(authorNames(i)) > 0 Then
            win.View.Reviewers.Item(authorNames(i)).Visible = visibleFlags(i)
        End If
        On Error GoTo Failed
    Next i

    ApplyReviewerFilter = True
    Exit Function

Failed:
    ApplyReviewerFilter = False
End Function

