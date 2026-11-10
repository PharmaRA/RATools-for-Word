Attribute VB_Name = "Macro_CaptionReferencePair"
Option Explicit

Private Type CaptionPairGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef value As CaptionPairGuid) As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef value As CaptionPairGuid) As Long
#End If
Private Const FEATURE_TITLE As String = "生成题注与引用"
Private Const SETTINGS_SECTION As String = "CaptionReferencePair"

Public Sub GenerateCaptionReferencePair(control As IRibbonControl)
    Wrapper_GenerateCaptionReferencePair
End Sub

Public Sub Wrapper_GenerateCaptionReferencePair()
    Dim panel As frmCaptionReference
    On Error GoTo Failed
    If Application.Documents.Count = 0 Then
        MsgBox "请先打开文档，将光标放在要生成题注的位置。", vbInformation, FEATURE_TITLE
        Exit Sub
    End If
    ValidateCaptionPairInsertionPoint Selection.Range
    Set panel = New frmCaptionReference
    panel.SetInsertionPoint Selection.Range.Duplicate
    panel.Show vbModal
CleanExit:
    If Not panel Is Nothing Then Unload panel
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, FEATURE_TITLE
    Resume CleanExit
End Sub

Public Sub ValidateCaptionPairInsertionPoint(ByVal target As Range)
    If target Is Nothing Then Err.Raise vbObjectError + 2400, FEATURE_TITLE, "请先定位插入光标。"
    If target.Start <> target.End Then
        Err.Raise vbObjectError + 2401, FEATURE_TITLE, "请先取消文字选区，将光标放在要生成题注的位置。"
    End If
    If target.Document.ReadOnly Or target.Document.ProtectionType <> wdNoProtection Then
        Err.Raise vbObjectError + 2402, FEATURE_TITLE, "当前文档只读或受保护，无法生成题注。"
    End If
    If target.StoryType <> wdMainTextStory Or target.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 2403, FEATURE_TITLE, "请将光标放在正文中、表格外的题注位置。"
    End If
    If target.Information(wdInFieldCode) Or target.Information(wdInFieldResult) Then
        Err.Raise vbObjectError + 2404, FEATURE_TITLE, "请将光标移到现有域之外，再生成题注。"
    End If
End Sub

Public Function NormalizeCaptionPairLabel(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) = 0 Then Err.Raise vbObjectError + 2405, FEATURE_TITLE, "请输入或选择题注标签。"
    ValidateSingleLine value, "题注标签"
    NormalizeCaptionPairLabel = value
End Function

Private Sub ValidateSingleLine(ByVal value As String, ByVal fieldName As String)
    Dim i As Long, ch As Long
    For i = 1 To Len(value)
        ch = AscW(Mid$(value, i, 1))
        If ch >= 0 And ch < 32 Then
            Err.Raise vbObjectError + 2406, FEATURE_TITLE, fieldName & "不能包含换行、制表符或控制字符。"
        End If
    Next i
End Sub

Public Function GetCaptionPairDefaultLabel() As String
    Dim savedLabel As String, label As CaptionLabel
    On Error Resume Next
    savedLabel = GetSetting("RATools", SETTINGS_SECTION, "LastLabel", vbNullString)
    On Error GoTo 0
    For Each label In Application.CaptionLabels
        If StrComp(label.Name, savedLabel, vbTextCompare) = 0 Then
            GetCaptionPairDefaultLabel = label.Name
            Exit Function
        End If
    Next label
    GetCaptionPairDefaultLabel = Application.CaptionLabels(wdCaptionTable).Name
End Function

Public Sub RememberCaptionPairLabel(ByVal labelName As String)
    ' 设置写入失败不能使已经完成的文档操作被误报为失败。
    On Error Resume Next
    SaveSetting "RATools", SETTINGS_SECTION, "LastLabel", labelName
    On Error GoTo 0
End Sub

Private Function ResolveCaptionPairLabel(ByVal labelName As String) As String
    Dim label As CaptionLabel
    For Each label In Application.CaptionLabels
        If StrComp(label.Name, labelName, vbTextCompare) = 0 Then
            ResolveCaptionPairLabel = label.Name
            Exit Function
        End If
    Next label
    Set label = Application.CaptionLabels.Add(Name:=labelName)
    ResolveCaptionPairLabel = label.Name
End Function

Private Function NextCaptionPairBookmark(ByVal doc As Document) As String
    Dim value As CaptionPairGuid, candidate As String, i As Long
    Do
        If CoCreateGuid(value) <> 0 Then
            Err.Raise vbObjectError + 2408, FEATURE_TITLE, "无法生成题注的唯一标识，请重试。"
        End If
        candidate = "_RefRA" & Right$("00000000" & Hex$(value.Data1), 8) & _
                    Right$("0000" & Hex$(value.Data2), 4) & Right$("0000" & Hex$(value.Data3), 4)
        For i = 0 To 7
            candidate = candidate & Right$("00" & Hex$(value.Data4(i)), 2)
        Next i
    Loop While doc.Bookmarks.Exists(candidate)
    ' 不复用已消失的书签名：文档中仍可能保留指向旧题注的失效引用。
    NextCaptionPairBookmark = candidate
End Function

' 在明确的插入点创建相邻两段；不依赖剪贴板，也不按题注文字猜测引用目标。
Public Function InsertCaptionReferencePair(ByVal insertionPoint As Range, _
                                          ByVal labelName As String, _
                                          ByVal captionText As String) As Range
    Dim doc As Document, point As Range, captionRange As Range, referenceRange As Range
    Dim bookmarkRange As Range, pairRange As Range, sequence As Field, item As Field
    Dim blockStart As Long, pairStart As Long, captionStart As Long, prefix As String, title As String
    Dim bookmarkName As String, undoStarted As Boolean, changed As Boolean
    Dim errorNumber As Long, errorText As String, rolledBack As Boolean

    ValidateCaptionPairInsertionPoint insertionPoint
    labelName = NormalizeCaptionPairLabel(labelName)
    captionText = Trim$(captionText)
    ValidateSingleLine captionText, "题注文字"
    Set doc = insertionPoint.Document
    labelName = ResolveCaptionPairLabel(labelName)
    Set point = insertionPoint.Duplicate
    blockStart = point.Start
    If point.Start > point.Paragraphs(1).Range.Start Then prefix = vbCr
    pairStart = blockStart + Len(prefix)
    captionStart = pairStart + 1
    If Len(captionText) > 0 Then title = " " & captionText

    On Error GoTo Failed
    Application.UndoRecord.StartCustomRecord FEATURE_TITLE
    undoStarted = True
    ' 两个空段依次容纳引用和题注；保留插入点两侧原有内容。
    point.Text = prefix & vbCr & vbCr
    changed = True
    Set point = doc.Range(captionStart, captionStart)
    point.InsertCaption Label:=labelName, Title:=title, _
                        Position:=wdCaptionPositionAbove, ExcludeLabel:=False
    Set captionRange = doc.Range(captionStart, captionStart).Paragraphs(1).Range
    For Each item In captionRange.Fields
        item.Update
        If item.Type = wdFieldSequence Then Set sequence = item
    Next item
    If sequence Is Nothing Then Err.Raise vbObjectError + 2407, FEATURE_TITLE, "未能创建题注编号。"

    ' 包含完整 SEQ 域结束符，保证剪切整段题注时书签随题注移动。
    Set bookmarkRange = doc.Range(captionRange.Start, sequence.Result.End + 1)
    bookmarkName = NextCaptionPairBookmark(doc)
    doc.Bookmarks.Add Name:=bookmarkName, Range:=bookmarkRange
    Set referenceRange = doc.Range(pairStart, pairStart)
    referenceRange.Style = wdStyleNormal
    referenceRange.InsertCrossReference ReferenceType:=wdRefTypeBookmark, _
                                        ReferenceKind:=wdContentText, ReferenceItem:=bookmarkName, _
                                        InsertAsHyperlink:=True, IncludePosition:=False
    Set referenceRange = doc.Range(pairStart, pairStart).Paragraphs(1).Range
    For Each item In referenceRange.Fields
        If item.Type = wdFieldRef And InStr(1, item.Code.Text, "MERGEFORMAT", vbTextCompare) = 0 Then
            item.Code.Text = item.Code.Text & " \* MERGEFORMAT "
        End If
        item.Update
        If item.Type = wdFieldRef Then item.Result.Font.Color = wdColorBlue
    Next item
    Set pairRange = doc.Range(pairStart, captionRange.End)
    Application.UndoRecord.EndCustomRecord
    undoStarted = False
    Set InsertCaptionReferencePair = pairRange
    Exit Function

Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If undoStarted Then Application.UndoRecord.EndCustomRecord
    If changed And undoStarted Then rolledBack = doc.Undo(1)
    On Error GoTo 0
    If changed And Not rolledBack Then errorText = errorText & vbCrLf & "自动撤销未成功，请检查本次插入位置并手动撤销。"
    Err.Raise errorNumber, FEATURE_TITLE, errorText
End Function
