Attribute VB_Name = "Macro_RemoveUnusedStyles"
Option Explicit

' =============================================
' 宏名称：RemoveUnusedStyles
' 功能：清理文件中未使用样式
' =============================================

Sub RemoveUnusedStyles()
    ' 定义变量
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim deleteCount As Integer
    Dim rng As Range
    
    ' 初始化
    Set doc = ActiveDocument
    deleteCount = 0
    
    ' 关闭屏幕更新以加快运行速度
    Application.ScreenUpdating = False
    
    ' 使用倒序循环遍历所有样式
    ' 倒序是因为在删除集合中的项目时，正序循环可能会导致跳过某些项目
    For i = doc.Styles.count To 1 Step -1
        Set sty = doc.Styles(i)
        
        ' 忽略错误（有些特殊样式无法被访问或删除，防止代码中断）
        On Error Resume Next
        
        ' 检查样式名称是否以 "-F" 结尾
        ' UCase 确保不区分大小写（例如 -f 和 -F 都会被检测）
        If HasStyleSuffix(sty.NameLocal) Then
            
            ' 确保只删除非内置样式（安全起见）
            If sty.BuiltIn = False Then
                
                ' 创建一个搜索范围，检查样式是否在文档中被实际使用
                If Not IsStyleUsedInDocument(doc, sty.NameLocal) Then
                    Err.Clear
                    sty.Delete
                    If Err.Number = 0 Then deleteCount = deleteCount + 1
                    Err.Clear
                End If
            End If
        End If
        
        ' 恢复正常的错误处理
        On Error GoTo 0
    Next i
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    ' 弹出结果
    MsgBox "操作完成！共删除了 " & deleteCount & " 个未使用的 '-F' 样式。", vbInformation, "清理完成"

End Sub


Private Function IsStyleUsedInDocument(ByVal doc As Document, ByVal styleName As String) As Boolean
    Dim root As Range
    Dim story As Range
    Dim dependent As Style
    Dim section As Section
    Dim headerFooter As HeaderFooter
    ' A failed inspection is not evidence that a style is unused.
    IsStyleUsedInDocument = True
    On Error GoTo UnknownUsage
    ' A base referenced by another style is not unused, even if that descendant
    ' is not currently applied. Preserve the definition rather than break links.
    For Each dependent In doc.Styles
        If StrComp(dependent.NameLocal, styleName, vbTextCompare) <> 0 Then
            If StrComp(CStr(dependent.BaseStyle), styleName, vbTextCompare) = 0 Then Exit Function
        End If
    Next dependent
    For Each root In doc.StoryRanges
        Set story = root
        Do While Not story Is Nothing
            If RangeUsesNamedStyle(story, styleName) Then Exit Function
            Set story = story.NextStoryRange
        Loop
    Next root
    If ShapesUseNamedStyle(doc.Shapes, styleName) Then Exit Function
    For Each section In doc.Sections
        For Each headerFooter In section.Headers
            If ShapesUseNamedStyle(headerFooter.Shapes, styleName) Then Exit Function
        Next headerFooter
        For Each headerFooter In section.Footers
            If ShapesUseNamedStyle(headerFooter.Shapes, styleName) Then Exit Function
        Next headerFooter
    Next section
    IsStyleUsedInDocument = False
UnknownUsage:
End Function


Private Function RangeUsesNamedStyle(ByVal story As Range, ByVal styleName As String) As Boolean
    Dim paragraph As Paragraph
    Dim table As Table
    Dim search As Range
    RangeUsesNamedStyle = True
    On Error GoTo UnknownUsage
            For Each paragraph In story.Paragraphs
                If StrComp(CStr(paragraph.Style), styleName, vbTextCompare) = 0 Then Exit Function
            Next paragraph
            For Each table In story.Tables
                If StrComp(CStr(table.Style), styleName, vbTextCompare) = 0 Then Exit Function
            Next table
            Set search = story.Duplicate
            With search.Find
                .ClearFormatting
                .Text = ""
                .Style = styleName
                .Format = True
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                If .Execute Then Exit Function
            End With
    RangeUsesNamedStyle = False
UnknownUsage:
End Function

Private Function ShapesUseNamedStyle(ByVal shapes As Object, ByVal styleName As String) As Boolean
    Dim shape As Shape
    ShapesUseNamedStyle = True
    On Error GoTo UnknownUsage
    For Each shape In shapes
        If shape.Type = msoGroup Then
            If ShapesUseNamedStyle(shape.GroupItems, styleName) Then Exit Function
        ElseIf shape.Type = msoCanvas Then
            If ShapesUseNamedStyle(shape.CanvasItems, styleName) Then Exit Function
        ElseIf shape.TextFrame.HasText <> 0 Then
            If RangeUsesNamedStyle(shape.TextFrame.TextRange, styleName) Then Exit Function
        End If
    Next shape
    ShapesUseNamedStyle = False
UnknownUsage:
End Function
