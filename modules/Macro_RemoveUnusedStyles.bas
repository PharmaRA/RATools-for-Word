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
                Set rng = doc.Content
                With rng.Find
                    .ClearFormatting
                    .Style = sty.NameLocal
                    .Format = True
                    .Execute FindText:="", Format:=True
                    
                    ' 如果 .Found 为 False，说明文档没用到这个样式
                    If .Found = False Then
                        sty.Delete
                        ' 如果删除成功（没有报错），计数加1
                        If Err.Number = 0 Then
                            deleteCount = deleteCount + 1
                        Else
                            ' 清除错误状态
                            Err.Clear
                        End If
                    End If
                End With
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
