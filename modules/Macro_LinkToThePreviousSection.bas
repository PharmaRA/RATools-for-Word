Attribute VB_Name = "Macro_LinkToThePreviousSection"
Option Explicit
Sub LinkToThePreviousSection()
    ' ==========================================
    ' 宏名称：LinkToThePreviousSection
    ' 描述：遍历文档中除第一节以外的所有节，
    '       将所有页眉和页脚设置为“链接到前一节”。
    ' ==========================================

    Dim sec As Section
    Dim hd As HeaderFooter
    Dim ft As HeaderFooter

    ' 如果文档只有一节，则不需要链接，直接退出
    If ActiveDocument.Sections.count < 2 Then
        MsgBox "文档只有一节，无需链接。", vbInformation, "提示"
        Exit Sub
    End If

    ' 开始遍历文档中的每一节
    For Each sec In ActiveDocument.Sections
        
        ' 跳过第一节（Index = 1），因为第一节无法链接到前一节
        If sec.Index > 1 Then
            
            ' 1. 处理该节的所有页眉
            ' (包括主页眉、首页页眉、偶数页页眉)
            For Each hd In sec.Headers
                ' 将“链接到前一条”属性设置为 True
                hd.LinkToPrevious = True
            Next hd
            
            ' 2. 处理该节的所有页脚
            ' (包括主页脚、首页页脚、偶数页页脚)
            For Each ft In sec.Footers
                ' 将“链接到前一条”属性设置为 True
                ft.LinkToPrevious = True
            Next ft
            
        End If
        
    Next sec

    ' 完成后通知用户
    MsgBox "处理完成！所有节的页眉和页脚已链接到前一节。", vbInformation, "成功"

End Sub
