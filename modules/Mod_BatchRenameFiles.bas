Attribute VB_Name = "Mod_BatchRenameFiles"
Option Explicit
Sub BatchRenameFiles()
    ' ==============================================================================
    ' 功能：批量修改文件名
    '       1. 仅保留汉字、字母、数字、中划线和下划线
    '       2. 所有字母统一转换为小写
    '       3. 字母/数字间的空格改为中划线 "-"
    '       4. 汉字与字符间的空格（以及其他剩余空格）直接删除
    '       5. 其他非法字符替换为中划线 "-"
    '       6. 自动递归处理子文件夹，被占用文件自动创建副本
    ' ==============================================================================

    Dim fDialog As FileDialog
    Dim mode As VbMsgBoxResult
    Dim fileList As Collection
    Dim vFile As Variant
    Dim targetFolder As String
    Dim fullPath As String
    Dim fileName As String, newFileName As String
    Dim baseName As String, extName As String
    Dim cleanName As String
    Dim regExClean As Object
    Dim regExSpace As Object
    Dim fso As Object
    Dim renamedCount As Long
    Dim copyCount As Long
    Dim i As Long
    Dim newPath As String
    Dim dupCounter As Long
    
    ' 1. 询问用户模式
    mode = MsgBox("请选择操作模式：" & vbCrLf & vbCrLf & _
                  "【是 (Yes)】 选择一个文件夹 (递归处理所有子文件夹)" & vbCrLf & _
                  "【否 (No)】  选择具体文件 (支持按住Ctrl或Shift多选)", _
                  vbYesNoCancel + vbQuestion, "选择模式")
    
    If mode = vbCancel Then Exit Sub
    
    Set fileList = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 2. 根据模式获取文件列表
    If mode = vbYes Then
        ' --- 文件夹模式 (递归) ---
        Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
        fDialog.Title = "请选择包含待处理文件的文件夹"
        
        If fDialog.Show = -1 Then
            targetFolder = fDialog.SelectedItems(1)
            RecursiveGetFiles fso.GetFolder(targetFolder), fileList
        Else
            Exit Sub
        End If
    Else
        ' --- 文件多选模式 ---
        Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
        fDialog.Title = "请选择需要清洗文件名的文件（可多选）"
        fDialog.AllowMultiSelect = True
        fDialog.Filters.Clear
        fDialog.Filters.Add "所有文件", "*.*"
        
        If fDialog.Show = -1 Then
            For i = 1 To fDialog.SelectedItems.count
                fileList.Add fDialog.SelectedItems(i)
            Next i
        Else
            Exit Sub
        End If
    End If
    
    If fileList.count = 0 Then
        MsgBox "没有找到可处理的文件。", vbExclamation
        Exit Sub
    End If
    
    ' 3. 初始化正则对象
    
    ' (A) 清理非法字符正则：仅保留 a-z, 0-9, -, _, 汉字
    Set regExClean = CreateObject("VBScript.RegExp")
    With regExClean
        .Global = True
        .IgnoreCase = True
        .Pattern = "[^a-z0-9\-\_" & ChrW(&H4E00) & "-" & ChrW(&H9FA5) & "]"
    End With
    
    ' (B) 空格处理正则：匹配 "字母或数字 + 空格 + 字母或数字" 的情况
    ' 用于将单词间的空格转为中划线
    Set regExSpace = CreateObject("VBScript.RegExp")
    With regExSpace
        .Global = True
        .IgnoreCase = True
        ' Lookahead断言：匹配一个字符和空格，且后面紧跟着另一个字符
        .Pattern = "([a-z0-9])\s+(?=[a-z0-9])"
    End With
    
    renamedCount = 0
    copyCount = 0
    Application.ScreenUpdating = False
    
    ' 4. 统一循环处理
    For Each vFile In fileList
        fullPath = CStr(vFile)
        
        targetFolder = fso.GetParentFolderName(fullPath) & "\"
        fileName = fso.GetFileName(fullPath)
        baseName = fso.GetBaseName(fullPath)
        extName = "." & fso.GetExtensionName(fullPath)
        If extName = "." Then extName = ""
        
        ' --- 核心处理逻辑开始 ---
        
        ' 步骤 1：转换为小写
        baseName = LCase(baseName)
        extName = LCase(extName)
        
        ' 步骤 2：处理空格
        ' 情况A：字母/数字 之间的空格 -> 替换为中划线 (例如: "file 01" -> "file-01")
        baseName = regExSpace.Replace(baseName, "$1-")
        
        ' 情况B：剩余的所有空格（包括汉字与字母间、汉字间） -> 直接删除 (例如: "测试 file" -> "测试file")
        baseName = Replace(baseName, " ", "")
        
        ' 步骤 3：清理非法字符（替换为中划线）
        cleanName = regExClean.Replace(baseName, "-")
        
        ' --- 核心处理逻辑结束 ---
        
        If Len(cleanName) = 0 Then cleanName = "renamed-file"
        
        newFileName = cleanName & extName
        
        If fileName <> newFileName Then
            newPath = targetFolder & newFileName
            
            ' 防重名
            dupCounter = 1
            Do While fso.FileExists(newPath)
                newFileName = cleanName & "_" & dupCounter & extName
                newPath = targetFolder & newFileName
                dupCounter = dupCounter + 1
            Loop
            
            ' 尝试重命名或复制
            On Error Resume Next
            Err.Clear
            Name fullPath As newPath
            
            If Err.Number = 0 Then
                renamedCount = renamedCount + 1
            Else
                Err.Clear
                fso.CopyFile fullPath, newPath
                If Err.Number = 0 Then
                    copyCount = copyCount + 1
                End If
            End If
            On Error GoTo 0
        End If
        
    Next vFile
    
    Application.ScreenUpdating = True
    
    ' 5. 结果提示
    MsgBox "处理完成！" & vbCrLf & _
           "直接重命名: " & renamedCount & " 个" & vbCrLf & _
           "创建副本(原文件被占用): " & copyCount & " 个", _
           vbInformation, "批量修改文件名"
    
    Set regExClean = Nothing
    Set regExSpace = Nothing
    Set fso = Nothing
    Set fDialog = Nothing
    Set fileList = Nothing

End Sub

' ==========================================
' 辅助过程：递归获取文件夹及子文件夹下的所有文件
' ==========================================
Private Sub RecursiveGetFiles(ByVal oFolder As Object, ByRef colFiles As Collection)
    Dim oFile As Object
    Dim oSubFolder As Object
    
    On Error Resume Next
    
    For Each oFile In oFolder.Files
        If Left(oFile.Name, 2) <> "~$" Then
            colFiles.Add oFile.Path
        End If
    Next oFile
    
    For Each oSubFolder In oFolder.SubFolders
        RecursiveGetFiles oSubFolder, colFiles
    Next oSubFolder
    
    On Error GoTo 0
End Sub

