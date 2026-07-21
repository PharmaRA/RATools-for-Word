Attribute VB_Name = "Mod_BatchDetectHighlights"
Option Explicit

' ==========================================
' 宏名称：BatchDetectHighlights
' 功能：批量检测文档是否包含高亮（突出显示颜色）
' ==========================================

' 定义全局 FSO 对象，避免重复创建
Private fso As Object

Sub BatchDetectHighlights()
    Dim userChoice As String
    Dim reportDoc As Document
    Dim doc As Document
    Dim targetFiles As Collection
    Dim filePath As Variant
    Dim fDialog As FileDialog
    Dim folderPath As String
    Dim startTime As Double
    Dim firstHighlight As Range
    Dim detectionError As String
    
    ' 初始化 FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. 获取用户选择的模式 (默认值为 1)
    userChoice = InputBox("请输入数字选择检测模式：" & vbCrLf & vbCrLf & _
                          "1 - 检测【当前打开】的文件" & vbCrLf & _
                          "2 - 选择【多个文件】进行批量检测" & vbCrLf & _
                          "3 - 选择【文件夹】（包含子文件夹）检测所有 Word 文件", _
                          "高亮检测工具", "1")
    
    If userChoice = "" Then Exit Sub ' 用户取消
    
    ' 初始化文件集合
    Set targetFiles = New Collection
    
    ' 2. 根据选择处理逻辑
    Select Case userChoice
        Case "1" ' 当前文件
            If Documents.count = 0 Then
                MsgBox "当前没有打开的文档！", vbExclamation
                Exit Sub
            End If
            
            If TryFindFirstHighlight(ActiveDocument, firstHighlight, detectionError) Then
                If MsgBox("当前文档【包含】突出显示颜色。是否跳转到第一处？", _
                          vbQuestion + vbYesNo, "检测到高亮内容") = vbYes Then
                    ActiveDocument.Activate
                    firstHighlight.Select
                    ActiveWindow.ScrollIntoView firstHighlight, True
                End If
            ElseIf Len(detectionError) > 0 Then
                MsgBox "检测失败：" & detectionError, vbExclamation
            Else
                MsgBox "当前文档【不包含】突出显示颜色。", vbInformation
            End If
            Exit Sub ' 模式1不需要生成报告，直接结束
            
        Case "2" ' 选择多个文件
            Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
            With fDialog
                .Title = "请选择要检测的 Word 文件"
                .Filters.Clear
                .Filters.Add "Word 文档", "*.docx; *.doc; *.docm"
                .AllowMultiSelect = True
                If .Show = -1 Then
                    For Each filePath In .SelectedItems
                        targetFiles.Add filePath
                    Next filePath
                Else
                    Exit Sub
                End If
            End With
            
        Case "3" ' 选择文件夹（包含子文件夹）
            Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
            With fDialog
                .Title = "请选择包含 Word 文件的文件夹"
                If .Show = -1 Then
                    folderPath = .SelectedItems(1)
                    ' 调用递归函数扫描文件夹
                    Call RecursiveScan(folderPath, targetFiles)
                Else
                    Exit Sub
                End If
            End With
            
        Case Else
            MsgBox "输入无效，请输入 1、2 或 3。", vbCritical
            Exit Sub
    End Select
    
    ' 3. 如果没有找到文件
    If targetFiles.count = 0 Then
        MsgBox "未找到需要处理的文件。", vbExclamation
        Exit Sub
    End If
    
    ' 4. 开始批量处理
    Application.ScreenUpdating = False ' 关闭屏幕更新（防止界面刷新）
    startTime = Timer
    
    ' 创建报告文档
    Set reportDoc = Documents.Add
    With reportDoc.Content
        .InsertAfter "突出显示颜色检测报告" & vbCrLf
        .InsertAfter "检测时间: " & Now & vbCrLf
        .InsertAfter "总计文件: " & targetFiles.count & vbCrLf & vbCrLf
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True
        .Font.Size = 14
        .InsertParagraphAfter
    End With
    
    ' 插入表格头
    Dim tbl As Table
    Set tbl = reportDoc.Tables.Add(Range:=reportDoc.Characters.Last, NumRows:=1, NumColumns:=4)
    tbl.Borders.Enable = True
    tbl.Cell(1, 1).Range.Text = "文件路径"
    tbl.Cell(1, 2).Range.Text = "检测结果"
    tbl.Cell(1, 3).Range.Text = "命中内容"
    tbl.Cell(1, 4).Range.Text = "定位"
    
    ' 循环处理文件
    Dim hasHighlight As Boolean
    Dim buttonRange As Range
    Dim openError As String
    
    For Each filePath In targetFiles
        ' 每个文件必须重置状态，避免错误时沿用上一文件的检测结果。
        Set doc = Nothing
        Set firstHighlight = Nothing
        hasHighlight = False
        openError = ""
        detectionError = ""

        ' 仅在打开文件时忽略错误，后续检测和报告逻辑不使用 Resume Next。
        On Error Resume Next
        Err.Clear
        Set doc = Documents.Open(fileName:=filePath, Visible:=False, ReadOnly:=True, AddToRecentFiles:=False)
        If Err.Number <> 0 Then openError = Err.Description
        Err.Clear
        On Error GoTo 0

        tbl.Rows.Add

        If doc Is Nothing Or Len(openError) > 0 Then
            tbl.Cell(tbl.Rows.count, 1).Range.Text = filePath
            tbl.Cell(tbl.Rows.count, 2).Range.Text = "无法打开文件"
            tbl.Cell(tbl.Rows.count, 2).Range.Font.Color = wdColorGray50
            If Len(openError) > 0 Then
                tbl.Cell(tbl.Rows.count, 3).Range.Text = openError
            Else
                tbl.Cell(tbl.Rows.count, 3).Range.Text = "未知错误"
            End If
            tbl.Cell(tbl.Rows.count, 4).Range.Text = "--"
        Else
            hasHighlight = TryFindFirstHighlight(doc, firstHighlight, detectionError)

            If Len(detectionError) > 0 Then
                tbl.Cell(tbl.Rows.count, 1).Range.Text = filePath
                tbl.Cell(tbl.Rows.count, 2).Range.Text = "检测失败"
                tbl.Cell(tbl.Rows.count, 2).Range.Font.Color = wdColorGray50
                tbl.Cell(tbl.Rows.count, 3).Range.Text = detectionError
                tbl.Cell(tbl.Rows.count, 4).Range.Text = "--"
            ElseIf hasHighlight Then
                ' 有高亮：添加超链接并显示红色结果
                reportDoc.Hyperlinks.Add Anchor:=tbl.Cell(tbl.Rows.count, 1).Range, _
                                         Address:=filePath, _
                                         TextToDisplay:=filePath
                                         
                tbl.Cell(tbl.Rows.count, 2).Range.Text = "包含高亮"
                tbl.Cell(tbl.Rows.count, 2).Range.Font.Color = wdColorRed
                tbl.Cell(tbl.Rows.count, 2).Range.Font.Bold = True
                tbl.Cell(tbl.Rows.count, 3).Range.Text = GetHighlightSummary(firstHighlight)
                tbl.Cell(tbl.Rows.count, 4).Range.Text = ""
                Set buttonRange = tbl.Cell(tbl.Rows.count, 4).Range.Duplicate
                buttonRange.MoveEnd Unit:=wdCharacter, count:=-1
                reportDoc.Fields.Add Range:=buttonRange, _
                                     Type:=wdFieldMacroButton, _
                                     Text:="JumpToHighlightFromReport 双击跳转", _
                                     PreserveFormatting:=False
            Else
                ' 无高亮：仅显示纯文本路径
                tbl.Cell(tbl.Rows.count, 1).Range.Text = filePath
                
                tbl.Cell(tbl.Rows.count, 2).Range.Text = "无"
                tbl.Cell(tbl.Rows.count, 2).Range.Font.Color = wdColorGreen
                tbl.Cell(tbl.Rows.count, 3).Range.Text = "--"
                tbl.Cell(tbl.Rows.count, 4).Range.Text = "--"
            End If
        End If

        If Not doc Is Nothing Then
            On Error Resume Next
            doc.Close SaveChanges:=wdDoNotSaveChanges
            On Error GoTo 0
        End If
    Next filePath
    
    ' 清理与结束
    Set fso = Nothing
    Application.ScreenUpdating = True
    tbl.AutoFitBehavior (wdAutoFitWindow)
    
    MsgBox "检测完成！共扫描 " & targetFiles.count & " 个文件。" & vbCrLf & "耗时 " & Format(Timer - startTime, "0.00") & " 秒。", vbInformation
End Sub

' 从批量检测报告的当前行读取文件路径，打开文档并定位到第一处高亮。
Sub JumpToHighlightFromReport()
    Dim reportCell As Cell
    Dim filePath As String
    Dim targetDoc As Document
    Dim openDoc As Document
    Dim firstHighlight As Range
    Dim detectionError As String

    On Error GoTo JumpError

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "请在检测报告中双击“跳转”按钮。", vbExclamation
        Exit Sub
    End If

    Set reportCell = Selection.Rows(1).Cells(1)
    filePath = reportCell.Range.Text
    filePath = Replace(filePath, Chr(13), "")
    filePath = Replace(filePath, Chr(7), "")
    filePath = Trim(filePath)

    If Len(filePath) = 0 Or Dir(filePath) = "" Then
        MsgBox "未找到对应文件：" & vbCrLf & filePath, vbExclamation
        Exit Sub
    End If

    For Each openDoc In Documents
        If StrComp(openDoc.FullName, filePath, vbTextCompare) = 0 Then
            Set targetDoc = openDoc
            Exit For
        End If
    Next openDoc

    If targetDoc Is Nothing Then
        Set targetDoc = Documents.Open(fileName:=filePath, AddToRecentFiles:=False)
    End If

    targetDoc.Activate
    If TryFindFirstHighlight(targetDoc, firstHighlight, detectionError) Then
        firstHighlight.Select
        ActiveWindow.ScrollIntoView firstHighlight, True
    ElseIf Len(detectionError) > 0 Then
        MsgBox "检测失败：" & detectionError, vbExclamation
    Else
        MsgBox "文档中已找不到高亮内容，可能在生成报告后被修改。", vbInformation
    End If
    Exit Sub

JumpError:
    MsgBox "跳转失败：" & Err.Description, vbExclamation
End Sub

' ==========================================
' 辅助过程：递归扫描文件夹
' ==========================================
Sub RecursiveScan(ByVal folderPath As String, ByRef fileCollection As Collection)
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim ext As String
    
    On Error Resume Next ' 防止权限错误导致中断
    Set folder = fso.GetFolder(folderPath)
    
    ' 遍历当前文件夹下的文件
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Path))
        ' 检查扩展名，并排除临时文件（~$开头）
        If (ext = "docx" Or ext = "doc" Or ext = "docm") And Left(file.Name, 2) <> "~$" Then
            fileCollection.Add file.Path
        End If
    Next file
    
    ' 递归遍历子文件夹
    For Each subFolder In folder.SubFolders
        Call RecursiveScan(subFolder.Path, fileCollection)
    Next subFolder
    On Error GoTo 0
End Sub

' ==========================================
' 辅助函数：检测单个文档是否包含高亮
' ==========================================
Function CheckForHighlight(doc As Document) As Boolean
    Dim firstHighlight As Range
    Dim detectionError As String

    CheckForHighlight = TryFindFirstHighlight(doc, firstHighlight, detectionError)
End Function

' 安全调用检测函数，并将检测异常与“没有高亮”区分开。
Function TryFindFirstHighlight(doc As Document, _
                               ByRef foundRange As Range, _
                               ByRef errorMessage As String) As Boolean
    On Error GoTo FindError

    Set foundRange = Nothing
    errorMessage = ""
    TryFindFirstHighlight = FindFirstHighlight(doc, foundRange)
    Exit Function

FindError:
    Set foundRange = Nothing
    errorMessage = Err.Description
    TryFindFirstHighlight = False
End Function

' 查找第一处含有可见内容的真实高亮。
' 纯空白、段落/单元格结束标记和隐藏文字不会被视为高亮内容。
Function FindFirstHighlight(doc As Document, ByRef foundRange As Range) As Boolean
    Dim rootStoryRange As Range
    Dim storyRange As Range
    Dim searchRange As Range
    Dim storyEnd As Long
    Dim scanStart As Long
    Dim nextStart As Long

    Set foundRange = Nothing
    FindFirstHighlight = False

    For Each rootStoryRange In doc.StoryRanges
        Set storyRange = rootStoryRange
        Do
            storyEnd = storyRange.End
            scanStart = storyRange.Start

            Do While scanStart < storyEnd
                Set searchRange = storyRange.Duplicate
                searchRange.SetRange Start:=scanStart, End:=storyEnd

                With searchRange.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    ' 逐字符查找，避免空文本格式查找在整段均为高亮时漏报。
                    .Text = "?"
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = True
                    .MatchWildcards = True
                    .Highlight = True
                End With

                If Not searchRange.Find.Execute Then Exit Do

                If IsVisibleHighlightCharacter(searchRange) Then
                    Set foundRange = ExpandVisibleHighlightRange(storyRange, searchRange)
                    FindFirstHighlight = True
                    Exit Function
                End If

                ' 当前命中只有不可见字符时，越过它继续查找本故事范围。
                nextStart = searchRange.End
                If nextStart <= scanStart Then nextStart = scanStart + 1
                If nextStart >= storyEnd Then Exit Do
                scanStart = nextStart
            Loop

            Set storyRange = storyRange.NextStoryRange
        Loop While Not storyRange Is Nothing
    Next rootStoryRange
End Function

' 从首个可见高亮字符开始，扩展到相邻的可见高亮字符，便于跳转和生成摘要。
Private Function ExpandVisibleHighlightRange(ByVal storyRange As Range, _
                                             ByVal firstCharacter As Range) As Range
    Dim characterRange As Range
    Dim visibleEnd As Long

    Set ExpandVisibleHighlightRange = firstCharacter.Duplicate
    visibleEnd = firstCharacter.End

    Do While visibleEnd < storyRange.End
        Set characterRange = storyRange.Duplicate
        characterRange.SetRange Start:=visibleEnd, End:=visibleEnd + 1
        If Not IsVisibleHighlightCharacter(characterRange) Then Exit Do
        visibleEnd = characterRange.End
    Loop

    ExpandVisibleHighlightRange.End = visibleEnd
End Function

' 判断字符是否是用户可见的高亮内容，而不是隐藏文字或格式控制字符。
Private Function IsVisibleHighlightCharacter(ByVal characterRange As Range) As Boolean
    Dim characterText As String
    Dim characterCode As Long

    If characterRange.Font.Hidden = True Then Exit Function
    If characterRange.HighlightColorIndex = wdNoHighlight Then Exit Function

    characterText = characterRange.Text
    If Len(characterText) = 0 Then Exit Function

    characterCode = AscW(Left$(characterText, 1))
    If characterCode < 0 Then characterCode = characterCode + 65536

    Select Case characterCode
        Case 0 To 32, 160, &H200B, &H200C, &H200D, &H2060, &HFEFF
            Exit Function
    End Select

    IsVisibleHighlightCharacter = True
End Function

' 生成报告中的故事范围类型和命中文本摘要。
Private Function GetHighlightSummary(ByVal highlightRange As Range) As String
    Dim summaryText As String

    If highlightRange Is Nothing Then
        GetHighlightSummary = "--"
        Exit Function
    End If

    summaryText = highlightRange.Text
    summaryText = Replace(summaryText, vbCr, " ")
    summaryText = Replace(summaryText, vbLf, " ")
    summaryText = Replace(summaryText, vbTab, " ")
    summaryText = Replace(summaryText, Chr$(7), " ")
    summaryText = Trim$(summaryText)
    If Len(summaryText) > 40 Then summaryText = Left$(summaryText, 40) & "..."

    GetHighlightSummary = GetStoryTypeName(highlightRange.StoryType) & "：" & summaryText
End Function

Private Function GetStoryTypeName(ByVal storyType As WdStoryType) As String
    Select Case storyType
        Case wdMainTextStory: GetStoryTypeName = "正文"
        Case wdFootnotesStory: GetStoryTypeName = "脚注"
        Case wdEndnotesStory: GetStoryTypeName = "尾注"
        Case wdCommentsStory: GetStoryTypeName = "批注"
        Case wdTextFrameStory: GetStoryTypeName = "文本框"
        Case wdEvenPagesHeaderStory, wdPrimaryHeaderStory, wdFirstPageHeaderStory
            GetStoryTypeName = "页眉"
        Case wdEvenPagesFooterStory, wdPrimaryFooterStory, wdFirstPageFooterStory
            GetStoryTypeName = "页脚"
        Case Else: GetStoryTypeName = "其他区域"
    End Select
End Function

