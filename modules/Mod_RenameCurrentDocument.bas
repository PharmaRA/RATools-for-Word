Attribute VB_Name = "Mod_RenameCurrentDocument"
Option Explicit
Sub RenameCurrentDocument()
    Dim strOldPath As String
    Dim strNewName As String
    Dim strPath As String
    Dim strExt As String
    Dim strNewFullPath As String
    
    ' 1. 检查文件是否已保存过 (如果是新建未保存的文档，无法重命名)
    If ActiveDocument.Path = "" Then
        MsgBox "请先保存一次文档，然后再使用此功能修改文件名。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 2. 获取当前文件信息
    strOldPath = ActiveDocument.FullName ' 完整的旧路径
    strPath = ActiveDocument.Path & Application.PathSeparator ' 文件夹路径
    
    ' 获取文件扩展名 (通过反向查找最后一个点号)
    strExt = "." & Split(ActiveDocument.Name, ".")(UBound(Split(ActiveDocument.Name, ".")))
    
    ' 3. 弹出输入框询问新文件名
    ' 默认显示当前文件名(不含扩展名)
    strNewName = InputBox("请输入新的文件名（不需要输入后缀名）：", "修改文件名", Replace(ActiveDocument.Name, strExt, ""))
    
    ' 4. 验证输入
    If Trim(strNewName) = "" Then Exit Sub ' 如果用户取消或未输入，则退出
    
    ' 构建新文件的完整路径
    strNewFullPath = strPath & strNewName & strExt
    
    ' 检查新文件名是否与旧文件名相同
    If strNewFullPath = strOldPath Then Exit Sub
    
    ' 检查新文件名是否已存在
    If Dir(strNewFullPath) <> "" Then
        MsgBox "该文件夹下已存在同名文件，请重新命名！", vbCritical, "错误"
        Exit Sub
    End If
    
    ' 5. 执行“重命名”操作 (另存为 + 删除旧文件)
    On Error GoTo ErrorHandler
    
    ' 另存为新文件
    ActiveDocument.SaveAs2 fileName:=strNewFullPath
    
    ' 删除旧文件
    Kill strOldPath
    
    ' 强制标记文档为“已保存”状态，防止关闭时提示保存
    ActiveDocument.Saved = True
    
    MsgBox "文件名已成功修改！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "发生错误，无法修改文件名。" & vbCrLf & "错误信息: " & Err.Description, vbCritical, "错误"
End Sub
