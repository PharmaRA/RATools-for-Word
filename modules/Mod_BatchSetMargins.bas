Attribute VB_Name = "Mod_BatchSetMargins"
Option Explicit

' ==========================================
' 主程序：运行此宏开始操作
' ==========================================
Sub BatchSetMargins()
    Dim choice As String
    Dim fd As FileDialog
    Dim selectedPath As String
    Dim startTime As Single
    Dim i As Long ' 用于循环计数
    
    ' 询问用户想要执行的操作模式
    choice = InputBox("请输入数字选择模式：" & vbCrLf & vbCrLf & _
                      "1 - 仅处理当前打开的文档" & vbCrLf & _
                      "2 - 选择文件进行处理（支持多选）" & vbCrLf & _
                      "3 - 选择文件夹批量处理所有Word文件", "选择模式", "1")
    
    If choice = "" Then Exit Sub ' 用户取消
    
    startTime = Timer
    ' 注意：如果不刷新屏幕，用户在模式1可能看不到即时变化，
    ' 所以我们在模式1时保持刷新，模式2/3关闭以提速
    If choice <> "1" Then Application.ScreenUpdating = False
    
    Select Case choice
        Case "1"
            ' 模式 1: 处理当前文档
            If Documents.count > 0 Then
                LogAction "正在处理当前文档..."
                ' 传递 True 表示需要弹出提示框告知结果
                ProcessSingleDoc ActiveDocument, True
            Else
                MsgBox "没有打开的文档！", vbExclamation
            End If
            
        Case "2"
            ' 模式 2: 选择文件 (支持多选)
            Set fd = Application.FileDialog(msoFileDialogFilePicker)
            With fd
                .Title = "请选择要处理的Word文件（按住Ctrl或Shift可多选）"
                .AllowMultiSelect = True ' *** 开启多选支持 ***
                .Filters.Clear
                .Filters.Add "Word 文件", "*.doc; *.docx; *.docm"
                
                If .Show = -1 Then
                    ' 遍历选中的每一个文件
                    For i = 1 To .SelectedItems.count
                        selectedPath = .SelectedItems(i)
                        ' 更新状态栏显示进度
                        Application.StatusBar = "正在处理 (" & i & "/" & .SelectedItems.count & "): " & Dir(selectedPath)
                        ProcessFile selectedPath
                    Next i
                    
                    Application.StatusBar = "" ' 清除状态栏
                    MsgBox "成功处理了 " & .SelectedItems.count & " 个文件！", vbInformation, "完成"
                End If
            End With
            
        Case "3"
            ' 模式 3: 文件夹批量处理
            Set fd = Application.FileDialog(msoFileDialogFolderPicker)
            With fd
                .Title = "请选择包含Word文件的文件夹"
                If .Show = -1 Then
                    selectedPath = .SelectedItems(1)
                    ProcessFolder selectedPath
                    MsgBox "文件夹批量处理完成！耗时: " & Format(Timer - startTime, "0.00") & "秒", vbInformation, "成功"
                End If
            End With
            
        Case Else
            MsgBox "无效的选择。", vbExclamation
    End Select
    
    Application.ScreenUpdating = True ' 确保最后恢复屏幕刷新
End Sub

' ==========================================
' 辅助过程：处理单个文件路径（打开-修改-保存-关闭）
' ==========================================
Private Sub ProcessFile(filePath As String)
    Dim doc As Document
    Dim isOpened As Boolean
    
    ' 检查文件是否已经打开
    On Error Resume Next
    Set doc = Documents(filePath) ' 尝试通过路径获取
    If Err.Number <> 0 Then
        ' 文件未打开，需要打开
        Err.Clear
        Set doc = Documents.Open(fileName:=filePath, Visible:=False)
        isOpened = True
    Else
        isOpened = False
    End If
    On Error GoTo 0
    
    If Not doc Is Nothing Then
        ' 调用处理逻辑，False 表示不弹出单个成功提示
        ProcessSingleDoc doc, False
        
        ' 如果是我们刚才打开的，则保存并关闭
        If isOpened Then
            doc.Save
            doc.Close
        End If
    End If
End Sub

' ==========================================
' 辅助过程：处理文档对象（检查保护 + 遍历节）
' ==========================================
Private Sub ProcessSingleDoc(doc As Document, showMsg As Boolean)
    ' 检查文档是否受保护（只读）
    If doc.ProtectionType <> wdNoProtection Then
        If showMsg Then MsgBox "当前文档受保护（只读），无法修改边距。", vbCritical, "失败"
        Exit Sub
    End If

    ' 核心修改逻辑
    Dim success As Boolean
    success = SetMarginsSafely(doc)
    
    If showMsg Then
        If success Then
            doc.Repaginate ' 强制重排以显示效果
            MsgBox "当前文档处理完成！所有节的边距已调整。", vbInformation, "成功"
        Else
            MsgBox "处理过程中遇到错误，请检查文档页面设置。", vbExclamation, "警告"
        End If
    End If
End Sub

' ==========================================
' 辅助过程：遍历文件夹
' ==========================================
Private Sub ProcessFolder(folderPath As String)
    Dim fileCollection As New Collection
    Dim i As Long

    ' 递归收集 Word 文档（doc/docx/docm，排除 ~$ 临时文件），与其他批处理宏行为一致
    CollectWordFiles folderPath, fileCollection

    If fileCollection.count = 0 Then
        Application.StatusBar = ""
        MsgBox "所选文件夹（含子文件夹）中未找到 Word 文档。", vbExclamation
        Exit Sub
    End If

    For i = 1 To fileCollection.count
        Application.StatusBar = "正在处理 (" & i & "/" & fileCollection.count & "): " & fileCollection(i)
        ProcessFile CStr(fileCollection(i))
    Next i

    Application.StatusBar = ""
End Sub

' 递归收集 Word 文档路径
Private Sub CollectWordFiles(ByVal sPath As String, ByRef fCollection As Collection)
    Dim FSO As Object, Folder As Object, SubFolder As Object, File As Object
    Dim ext As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set Folder = FSO.GetFolder(sPath)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    For Each File In Folder.Files
        ext = LCase(FSO.GetExtensionName(File.Name))
        If (ext = "doc" Or ext = "docx" Or ext = "docm") Then
            If Left(File.Name, 2) <> "~$" Then fCollection.Add File.Path
        End If
    Next

    For Each SubFolder In Folder.SubFolders
        CollectWordFiles SubFolder.Path, fCollection
    Next
    Set FSO = Nothing
End Sub

' ==========================================
' 核心逻辑：安全设置边距（防 4608 错误版 + 遍历 Sections）
' ==========================================
Private Function SetMarginsSafely(doc As Document) As Boolean
    Dim targetPoints As Single
    Dim sec As Section
    Dim hasError As Boolean
    
    ' 2.54 厘米 = 72 磅
    targetPoints = 72
    hasError = False
    
    ' 开启容错模式
    On Error Resume Next
    
    ' Word 文档可能包含多个节，直接设置 doc.PageSetup 有时只对第一节生效
    For Each sec In doc.Sections
        Err.Clear
        
        ' --- 尝试 1: 标准转换 ---
        With sec.PageSetup
            .TopMargin = Application.CentimetersToPoints(2.54)
            .BottomMargin = Application.CentimetersToPoints(2.54)
            .LeftMargin = Application.CentimetersToPoints(2.54)
            .RightMargin = Application.CentimetersToPoints(2.54)
        End With
        
        ' --- 检查错误 ---
        If Err.Number <> 0 Then
            Err.Clear ' 清除错误，尝试备用方案
            
            ' --- 尝试 2: 直接磅值 (绕过转换函数错误) ---
            With sec.PageSetup
                .TopMargin = targetPoints
                .BottomMargin = targetPoints
                .LeftMargin = targetPoints
                .RightMargin = targetPoints
            End With
            
            ' 如果依然失败
            If Err.Number <> 0 Then
                Debug.Print "节处理失败: " & doc.Name & " Section:" & sec.Index
                hasError = True
            End If
        End If
    Next sec
    
    ' 恢复错误处理
    On Error GoTo 0
    
    SetMarginsSafely = Not hasError
End Function

Private Sub LogAction(msg As String)
    ' 简单的状态栏日志
    Application.StatusBar = msg
End Sub

