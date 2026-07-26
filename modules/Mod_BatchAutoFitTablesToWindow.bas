Attribute VB_Name = "Mod_BatchAutoFitTablesToWindow"
Sub BatchAutoFitTablesToWindow()
    '================================================
    ' 功能：将文档中所有表格批量设置为“根据窗口自动调整”
    '================================================
    Dim objTable As Table
    Dim count As Integer
    count = 0
    
    ' 关闭屏幕更新，加快处理速度，防止屏幕闪烁
    Application.ScreenUpdating = False
    
    ' 检查文档中是否有表格
    If ActiveDocument.Tables.count > 0 Then
        ' 循环遍历每一个表格
        For Each objTable In ActiveDocument.Tables
            ' 应用“根据窗口自动调整”
            objTable.AutoFitBehavior (wdAutoFitWindow)
            count = count + 1
        Next objTable
    Else
        Application.ScreenUpdating = True
        MsgBox "当前文档中没有发现表格。", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    ' 弹出完成提示
    MsgBox "处理完成！已成功调整 " & count & " 个表格。", vbInformation, "成功"

End Sub
