Attribute VB_Name = "Mod_HyperlinksToBlue"
'=== 智能设置超链接和域为蓝色===
Sub SetHyperlinksAndFieldsToBlue()
    On Error GoTo ErrH
    Dim hyperlink As hyperlink
    Dim field As field
    Dim storyRange As Range
    Dim countChanged As Integer
    
    Application.ScreenUpdating = False
    countChanged = 0
    
    ' 1. 处理超链接
    For Each hyperlink In ActiveDocument.Hyperlinks
        If hyperlink.Range.Font.Color <> RGB(0, 0, 255) Then
            hyperlink.Range.Font.Color = RGB(0, 0, 255)
            countChanged = countChanged + 1
        End If
    Next hyperlink
    
    ' 2. 处理域，但排除题注和页码
    For Each storyRange In ActiveDocument.StoryRanges
        countChanged = countChanged + ProcessFieldsExcludeCaptions(storyRange)
    Next storyRange
    
    Application.ScreenUpdating = True
    
    If countChanged > 0 Then
        MsgBox "已将 " & countChanged & " 个超链接和域设置为蓝色", vbInformation
    Else
        MsgBox "所有超链接和域已经是蓝色", vbInformation
    End If
    Exit Sub

ErrH:
    Application.ScreenUpdating = True
    MsgBox "设置蓝色失败：" & Err.Description, vbCritical
End Sub

Private Function ProcessFieldsExcludeCaptions(rng As Range) As Integer
    Dim field As field
    Dim count As Integer
    Dim fieldCode As String
    
    count = 0
    
    ' 处理主范围
    For Each field In rng.Fields
        If IsProcessableField(field) Then
             If field.Result.Font.Color <> RGB(0, 0, 255) Then
                field.Result.Font.Color = RGB(0, 0, 255)
                count = count + 1
            End If
        End If
    Next field
    
    ' 处理链接的范围（如页眉页脚）
    Do While Not (rng.NextStoryRange Is Nothing)
        Set rng = rng.NextStoryRange
        For Each field In rng.Fields
            If IsProcessableField(field) Then
                If field.Result.Font.Color <> RGB(0, 0, 255) Then
                    field.Result.Font.Color = RGB(0, 0, 255)
                    count = count + 1
                End If
            End If
        Next field
    Loop
    
    ProcessFieldsExcludeCaptions = count
End Function

' 辅助函数：判断是否为需要处理的域（排除页码和题注）
Private Function IsProcessableField(fld As field) As Boolean
    ' 1. 排除页码相关域类型
    If fld.Type = wdFieldPage Or fld.Type = wdFieldNumPages Or fld.Type = wdFieldSectionPages Then
        IsProcessableField = False
        Exit Function
    End If
    
    ' 2. 排除题注（基于域代码文本）
    Dim fieldCode As String
    fieldCode = LCase(fld.Code.Text)
    
    If IsCaptionField(fieldCode) Then
        IsProcessableField = False
        Exit Function
    End If
    
    ' 通过检查
    IsProcessableField = True
End Function

' 判断是否为题注域的函数
Private Function IsCaptionField(fieldCode As String) As Boolean
    Dim captionIndicators As Variant
    captionIndicators = Array("seq", "图", "表", "chart", "figure", "table", "caption")
    
    Dim indicator As Variant
    For Each indicator In captionIndicators
        If InStr(1, fieldCode, indicator, vbTextCompare) > 0 Then
            IsCaptionField = True
            Exit Function
        End If
    Next indicator
    
    IsCaptionField = False
End Function

Public Sub SetHyperlinksAndFieldsToBlueRibbon(control As IRibbonControl)
    SetHyperlinksAndFieldsToBlue
End Sub
