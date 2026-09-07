Attribute VB_Name = "Macro_HyperlinksToBlue"
Option Explicit
'=== 智能设置超链接和域为蓝色===
Sub SetHyperlinksAndFieldsToBlue()
    On Error GoTo ErrH
    Dim hyperlink As hyperlink
    Dim field As field
    Dim storyRange As Range
    Dim countChanged As Long

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

Private Function ProcessFieldsExcludeCaptions(rng As Range) As Long
    Dim field As field
    Dim fieldCount As Long
    Dim fieldCode As String

    fieldCount = 0

    ' 处理主范围
    For Each field In rng.Fields
        If IsProcessableField(field) Then
             If field.Result.Font.Color <> RGB(0, 0, 255) Then
                field.Result.Font.Color = RGB(0, 0, 255)
                fieldCount = fieldCount + 1
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
                    fieldCount = fieldCount + 1
                End If
            End If
        Next field
    Loop

    ProcessFieldsExcludeCaptions = fieldCount
End Function

' 辅助函数：判断是否为需要处理的域（排除页码和题注）
Private Function IsProcessableField(fld As field) As Boolean
    ' 1. 排除页码相关域类型
    If fld.Type = wdFieldPage Or fld.Type = wdFieldNumPages Or fld.Type = wdFieldSectionPages Then
        IsProcessableField = False
        Exit Function
    End If

    ' 2. 排除题注序号域 (wdFieldSequence = 12) 与样式引用域 (wdFieldStyleRef = 10)
    ' 题注章节编号（如表3.2.2.6-1中的3.2.2.6）采用 STYLEREF 域，不属于交叉引用/超链接，不应标蓝
    If fld.Type = wdFieldSequence Or fld.Type = wdFieldStyleRef Then
        IsProcessableField = False
        Exit Function
    End If

    ' 3. 排除目录域 (wdFieldTOC = 13)
    If fld.Type = wdFieldTOC Then
        IsProcessableField = False
        Exit Function
    End If

    ' 4. 排除题注段落中的任何域（即使有未识别类型的域）
    If IsInCaptionParagraph(fld) Then
        IsProcessableField = False
        Exit Function
    End If

    ' 5. 排除题注相关文本的域代码（双重保险）
    Dim fieldCode As String
    fieldCode = LCase$(Trim$(fld.Code.Text))

    If IsCaptionField(fieldCode) Then
        IsProcessableField = False
        Exit Function
    End If

    ' 通过检查
    IsProcessableField = True
End Function

' 判断域是否位于题注段落中
Private Function IsInCaptionParagraph(fld As field) As Boolean
    On Error Resume Next
    Dim pStyleName As String
    pStyleName = LCase$(fld.Result.Paragraphs(1).Style.NameLocal)

    If InStr(1, pStyleName, "caption", vbTextCompare) > 0 Or _
       InStr(1, pStyleName, "题注", vbTextCompare) > 0 Or _
       InStr(1, pStyleName, "表标题", vbTextCompare) > 0 Or _
       InStr(1, pStyleName, "图标题", vbTextCompare) > 0 Or _
       InStr(1, pStyleName, "table title", vbTextCompare) > 0 Or _
       InStr(1, pStyleName, "figure title", vbTextCompare) > 0 Then
        IsInCaptionParagraph = True
        Exit Function
    End If
    On Error GoTo 0

    IsInCaptionParagraph = False
End Function

' 判断是否为题注域的函数
Private Function IsCaptionField(fieldCode As String) As Boolean
    Dim captionIndicators As Variant
    captionIndicators = Array("seq", "styleref", "图", "表", "chart", "figure", "table", "caption")

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
