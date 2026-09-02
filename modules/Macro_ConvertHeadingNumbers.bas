Attribute VB_Name = "Macro_ConvertHeadingNumbers"
Option Explicit
Sub ConvertHeadingNumbers()
    On Error GoTo ErrH
    ' 声明变量
    Dim i As Long
    Dim doc As Document
    Dim para As Paragraph
    Dim convertedCount As Long
    
    Set doc = ActiveDocument
    convertedCount = 0
    
    ' 关闭屏幕更新以加快速度
    Application.ScreenUpdating = False
    
    ' 【关键点】从最后一段开始，倒着向前循环 (Step -1)
    ' 这样转换后面的编号时，前面的编号结构还保留着，不会导致乱码
    For i = doc.Paragraphs.count To 1 Step -1
        Set para = doc.Paragraphs(i)
        
        ' 检查段落的大纲级别 (1-9 代表是标题)
        If para.Format.OutlineLevel <> wdOutlineLevelBodyText Then
            ' 检查该标题是否有编号
            If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                ' 将编号转换为静态文本
                para.Range.ListFormat.ConvertNumbersToText
                convertedCount = convertedCount + 1
            End If
        End If
    Next i
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    ' 提示完成
    MsgBox "修复版转换完成！共处理了 " & convertedCount & " 个标题。", vbInformation, "操作成功"
    Exit Sub

ErrH:
    Application.ScreenUpdating = True
    MsgBox "标题编号转换失败：" & Err.Description, vbCritical
End Sub
