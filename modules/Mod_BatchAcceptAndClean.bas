Attribute VB_Name = "Mod_BatchAcceptAndClean"
Option Explicit

' =============================================
' 宏名称：BatchAcceptAndClean
' 功能：Word文档批量清理工具
' =============================================

' === 主程序入口 ===
Sub BatchAcceptAndClean()
    Dim strMode As String
    Dim folderPath As String
    Dim fileCollection As New Collection
    Dim fileItem As Variant
    Dim i As Long
    Dim processedCount As Long
    Dim failedCount As Long
    
    ' 1. 模式选择
    strMode = InputBox("请输入模式编号：" & vbCrLf & vbCrLf & _
                       "1 - 【当前文档】处理当前打开的文档" & vbCrLf & _
                       "2 - 【文件模式】选择单个或多个文件" & vbCrLf & _
                       "3 - 【文件夹模式】递归处理文件夹", _
                       "Word批量清理工具", "1")
    
    If StrPtr(strMode) = 0 Or strMode = "" Then Exit Sub
    
    ' 2. 性能设置
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    On Error GoTo ErrorHandler
    processedCount = 0
    failedCount = 0
    
    Select Case strMode
        Case "1"
            If Documents.count > 0 Then
                Call DeepCleanDocument(ActiveDocument)
                processedCount = 1
                MsgBox "当前文档处理完成！", vbInformation
            End If
            
        Case "2"
            With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "选择Word文档"
                .AllowMultiSelect = True
                .Filters.Clear
                .Filters.Add "Word文档", "*.doc; *.docx; *.docm", 1
                If .Show = -1 Then
                    For Each fileItem In .SelectedItems
                        If ProcessFile(CStr(fileItem)) Then
                            processedCount = processedCount + 1
                        Else
                            failedCount = failedCount + 1
                        End If
                    Next
                End If
            End With
            
        Case "3"
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "选择根文件夹"
                If .Show = -1 Then
                    folderPath = .SelectedItems(1)
                    Application.StatusBar = "正在扫描文件..."
                    Call RecursiveFindFiles(folderPath, fileCollection)
                    
                    If fileCollection.count > 0 Then
                        For i = 1 To fileCollection.count
                            Application.StatusBar = "正在处理 [" & i & "/" & fileCollection.count & "]"
                            If ProcessFile(fileCollection(i)) Then
                                processedCount = processedCount + 1
                            Else
                                failedCount = failedCount + 1
                            End If
                        Next
                    Else
                        MsgBox "未找到Word文档", vbExclamation
                    End If
                End If
            End With
            
        Case Else
            MsgBox "无效输入", vbExclamation
    End Select

ExitHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    If failedCount > 0 And strMode <> "1" Then
        MsgBox "处理完成：成功 " & processedCount & " 个，失败 " & failedCount & _
               " 个（无法打开，请检查文件是否被占用或有密码）。", vbExclamation
    ElseIf processedCount > 0 And strMode <> "1" Then
        MsgBox "处理完成！共处理 " & processedCount & " 个文件。", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

' === 模块：处理单个文件 (后台模式 Visible=False) ===
Private Function ProcessFile(filePath As String) As Boolean
    Dim doc As Document
    On Error Resume Next
    
    ' 后台模式：Visible:=False 提升速度
    Set doc = Documents.Open(fileName:=filePath, Visible:=False, AddToRecentFiles:=False)
    
    If Err.Number = 0 Then
        Call DeepCleanDocument(doc)
        doc.Save
        doc.Close
        ProcessFile = True
    Else
        Debug.Print "打开失败: " & filePath
        ProcessFile = False
    End If
    On Error GoTo 0
End Function

' === 模块：清理逻辑 ===
Private Sub DeepCleanDocument(doc As Document)
    Dim rng As Range
    Dim storyIndex As Variant
    Dim commentRetry As Long ' 批注删除重试计数器
    ' 仅遍历核心区域，跳过无效区域
    ' 1=MainText, 2=Footnotes, 3=Endnotes, 5=Comments(虽删但需查), 6-11=Headers/Footers
    Dim targetStories As Variant
    targetStories = Array(1, 2, 3, 5, 6, 7, 8, 9, 10, 11)
    
    With doc
        ' 1. 解除文档保护 (仅当被保护时才执行)
        If .ProtectionType <> wdNoProtection Then .Unprotect

        ' 2. 全局快速清理 (处理大部分简单修订)
        On Error Resume Next
        .Revisions.AcceptAll
        On Error GoTo 0

        ' 强力循环删除所有批注 (解决带有答复的"现代批注"需要逐层剥离的问题)
        commentRetry = 0
        Do While .Comments.count > 0 And commentRetry < 20
            On Error Resume Next
            .DeleteAllComments
            On Error GoTo 0
            commentRetry = commentRetry + 1
        Loop

        ' 3. 链式遍历核心 Story (移除 Sections 循环，大幅提速)
        For Each storyIndex In targetStories
            On Error Resume Next
            Set rng = .StoryRanges(storyIndex)
            On Error GoTo 0
            
            ' NextStoryRange 会自动跳转到下一节的同类区域
            Do While Not rng Is Nothing
                Call CleanSingleRange(rng)
                Set rng = rng.NextStoryRange
            Loop
        Next storyIndex
        
        ' 4. 确保嵌套极深的残留批注被彻底删除
        commentRetry = 0
        Do While .Comments.count > 0 And commentRetry < 20
            On Error Resume Next
            .DeleteAllComments
            On Error GoTo 0
            commentRetry = commentRetry + 1
        Loop
        
        ' 5. 结束设置
        .TrackRevisions = False
    End With
End Sub

' === 模块：清理单个 Range (按需执行) ===
Private Sub CleanSingleRange(rng As Range)
    Dim shp As Shape
    Dim frm As Frame
    
    On Error Resume Next
    
    ' A. 智能解锁域 (仅当存在域时才操作)
    If rng.Fields.count > 0 Then rng.Fields.Locked = False
    
    ' B. 接受修订
    rng.Revisions.AcceptAll
    
    ' C. 递归清理 Shapes (仅当存在Shape时才进入循环)
    If rng.ShapeRange.count > 0 Then
        For Each shp In rng.ShapeRange
            Call ProcessShapeRecursively(shp)
        Next
    End If
    
    ' D. 清理 Frames (仅当存在Frame时才进入循环)
    If rng.Frames.count > 0 Then
        For Each frm In rng.Frames
            If frm.Range.Fields.count > 0 Then frm.Range.Fields.Locked = False
            frm.Range.Revisions.AcceptAll
        Next
    End If
    
    On Error GoTo 0
End Sub

' === 模块：递归处理图形对象 ===
Private Sub ProcessShapeRecursively(shp As Shape)
    Dim subShp As Shape
    
    On Error Resume Next
    
    ' 1. 文本框内容
    If Not shp.TextFrame Is Nothing Then
        If shp.TextFrame.HasText Then
            ' 仅解锁和接受
            If shp.TextFrame.TextRange.Fields.count > 0 Then shp.TextFrame.TextRange.Fields.Locked = False
            shp.TextFrame.TextRange.Revisions.AcceptAll
        End If
    End If
    
    ' 2. 组合图形 (Group)
    If shp.Type = msoGroup Then
        For Each subShp In shp.GroupItems
            Call ProcessShapeRecursively(subShp)
        Next
    End If
    
    ' 3. 画布 (Canvas)
    If shp.Type = msoCanvas Then
        For Each subShp In shp.CanvasItems
            Call ProcessShapeRecursively(subShp)
        Next
    End If
    
    On Error GoTo 0
End Sub

' === 模块：文件递归搜索 ===
Private Sub RecursiveFindFiles(ByVal sPath As String, ByRef fCollection As Collection)
    Dim FSO As Object, Folder As Object, SubFolder As Object, File As Object
    Dim ext As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set Folder = FSO.GetFolder(sPath)
    If Err.Number <> 0 Then Exit Sub
    
    ' 遍历文件
    For Each File In Folder.Files
        ext = LCase(FSO.GetExtensionName(File.Name))
        If (ext = "doc" Or ext = "docx" Or ext = "docm") Then
            If Left(File.Name, 2) <> "~$" Then fCollection.Add File.Path
        End If
    Next
    
    ' 递归子文件夹
    For Each SubFolder In Folder.SubFolders
        RecursiveFindFiles SubFolder.Path, fCollection
    Next
    Set FSO = Nothing
End Sub

