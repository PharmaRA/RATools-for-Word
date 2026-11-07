Attribute VB_Name = "Macro_PasteRelinkCrossReference"
Option Explicit

' =============================================
' 模块名：Macro_PasteRelinkCrossReference
' 功能：一步完成「粘贴 -> 重连交叉引用 -> 更新域」。
'
' 背景：同一文档内复制粘贴时，Word 不会把隐藏书签（_Ref########）
'       一起复制过去，粘贴出来的 REF 域仍指向原题注；而题注里的
'       SEQ 域会重新编号（表 1 -> 表 2），于是正文与题注对不上。
'       本模块在粘贴后，把粘贴范围内的 REF/PAGEREF 重新指向
'       粘贴出来的新题注书签，并且只更新粘贴范围内的域。
'
' 取舍：只处理落在「题注段落」（含 SEQ 域）里的书签；
'       指向标题、普通书签的引用保持原样。
' =============================================

Private Const DIALOG_TITLE As String = "粘贴并重连引用"
Private Const UNDO_RECORD_NAME As String = "粘贴并重连交叉引用"

' ====== TEST SUPPORT：静默模式与最后一次错误，供自动化测试驱动 ======
Private mQuietMode As Boolean
Private mLastErrorMessage As String

Public Sub SetPasteRelinkQuietMode(ByVal enabled As Boolean)
    mQuietMode = enabled
    mLastErrorMessage = vbNullString
End Sub

' 无参入口：Application.Run 传参需要 ByRef，测试里改用这个
Public Sub EnablePasteRelinkQuietMode()
    SetPasteRelinkQuietMode True
End Sub

Public Function GetLastPasteRelinkErrorForTest() As String
    GetLastPasteRelinkErrorForTest = mLastErrorMessage
End Function


' Ribbon 回调
Public Sub PasteRelinkCrossReference(control As IRibbonControl)
    Core_PasteRelinkCrossReference
End Sub

' 宏列表入口：Application.Run 不支持带参数的过程
Public Sub Wrapper_PasteRelinkCrossReference()
    Core_PasteRelinkCrossReference
End Sub

Public Sub Core_PasteRelinkCrossReference()
    Dim doc As Document
    Dim sel As Selection
    Dim pasteStart As Long
    Dim pasteEnd As Long
    Dim pastedRange As Range
    Dim pastedFields As Collection
    Dim captions As Collection
    Dim screenUpdating As Boolean
    Dim undoStarted As Boolean
    Dim ambiguousCaptions As Boolean

    If Application.Documents.Count = 0 Then
        ReportError "当前没有打开的文档。"
        Exit Sub
    End If

    mLastErrorMessage = vbNullString
    Set doc = ActiveDocument
    Set sel = Selection

    ' 页眉、脚注等故事区没有题注场景，退化为普通粘贴
    If sel.StoryType <> wdMainTextStory Then
        On Error Resume Next
        sel.PasteAndFormat wdFormatOriginalFormatting
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    screenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    pasteStart = sel.Start
    undoStarted = StartUndoRecord()

    On Error GoTo PasteFailed
    sel.PasteAndFormat wdFormatOriginalFormatting

    On Error GoTo RelinkFailed
    pasteEnd = sel.End

    If pasteEnd > pasteStart Then
        Set pastedRange = doc.Range(pasteStart, pasteEnd)
        ' 先抓取域对象：后面改写域代码可能改变文字长度，届时 Range 边界会失效
        Set pastedFields = CollectPastedFields(pastedRange)
        Set captions = CollectPastedCaptions(doc, pastedRange)
        If captions.Count > 0 Then
            RelinkPastedReferences doc, pastedRange, captions, pastedFields, ambiguousCaptions
        End If

        ' 单个域更新失败不影响整体：内容已经粘贴好了
        UpdateCollectedFields pastedFields
    End If

    On Error GoTo 0
    FinishUndoRecord undoStarted
    Application.ScreenUpdating = screenUpdating
    If ambiguousCaptions Then ReportError "存在标签和描述相同的多个题注，相关引用保留原目标。请手动确认这些引用。"
    Exit Sub

PasteFailed:
    FinishUndoRecord undoStarted
    Application.ScreenUpdating = screenUpdating
    ReportError "无法粘贴剪贴板内容。" & vbCrLf & vbCrLf & _
                "请先复制需要粘贴的内容，再点击本按钮。" & vbCrLf & _
                "错误信息：" & Err.Description
    Exit Sub

RelinkFailed:
    ' 粘贴已经完成，只是重连/更新阶段出错：保留已粘贴内容，按一次 Ctrl+Z 可整体撤销
    FinishUndoRecord undoStarted
    Application.ScreenUpdating = screenUpdating
    ReportError "内容已粘贴，但重连交叉引用时出错。" & vbCrLf & vbCrLf & _
                "粘贴的内容和原文都保持完整，可按一次 Ctrl+Z 整体撤销。" & vbCrLf & _
                "错误信息：" & Err.Description
End Sub

' 测试时静默，正常运行弹提示
Private Sub ReportError(ByVal message As String)
    mLastErrorMessage = message
    If mQuietMode Then Exit Sub
    MsgBox message, vbExclamation, DIALOG_TITLE
End Sub


' ==================== UndoRecord ====================

' 把粘贴与后续修复合并为一步，Ctrl+Z 可整体撤销
Private Function StartUndoRecord() As Boolean
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord UNDO_RECORD_NAME
    StartUndoRecord = (Err.Number = 0)
    Err.Clear
End Function

Private Sub FinishUndoRecord(ByVal started As Boolean)
    If Not started Then Exit Sub
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    Err.Clear
End Sub


' ==================== 粘贴块内的域 ====================

' 收集粘贴块内的全部域对象。Word 的 Fields 集合每次访问都会重算
' 范围，改写域代码后再按范围取域可能漏掉末尾的域，因此统一先抓对象。
Private Function CollectPastedFields(ByVal pastedRange As Range) As Collection
    Dim result As Collection
    Dim fld As Field

    Set result = New Collection
    On Error Resume Next
    For Each fld In pastedRange.Fields
        If IsRangeInside(fld.Code.Start, fld.Result.End, pastedRange.Start, pastedRange.End) Then
            result.Add fld
        End If
    Next fld
    Err.Clear
    Set CollectPastedFields = result
End Function

' 逐域更新；单个域失败（如引用已失效）不影响其余域。
' 必须分两轮：先更新 SEQ 让题注重新编号，再更新 REF/PAGEREF，
' 否则引用会读到尚未刷新的旧编号（显示「表 1」而题注已是「表 2」）。
Private Sub UpdateCollectedFields(ByVal fields As Collection)
    Dim fld As Field
    Dim pass As Long

    If fields Is Nothing Then Exit Sub

    On Error Resume Next

    ' 第一轮：除引用以外的域（SEQ、STYLEREF 等）
    For Each fld In fields
        If fld.Type <> wdFieldRef And fld.Type <> wdFieldPageRef Then
            fld.Update
            Err.Clear
        End If
    Next fld

    ' 第二轮：引用域，此时题注编号已是最新
    For Each fld In fields
        If fld.Type = wdFieldRef Or fld.Type = wdFieldPageRef Then
            fld.Update
            Err.Clear
        End If
    Next fld
End Sub


' ==================== 收集粘贴出来的题注 ====================

' 返回 Collection，每项是一个 Variant 数组：
' 0=标签名 1=描述文字 2=段落起点 3=段落终点 4=SEQ 结果起点 5=SEQ 结果终点
Private Function CollectPastedCaptions(ByVal doc As Document, ByVal pastedRange As Range) As Collection
    Dim result As Collection
    Dim fld As Field
    Dim para As Paragraph
    Dim pathLabel As String
    Dim pathDesc As String
    Dim captionStart As Long
    Dim captionEnd As Long

    Set result = New Collection

    For Each fld In pastedRange.Fields
        If fld.Type = wdFieldSequence Then
            If IsRangeInside(fld.Code.Start, fld.Result.End, pastedRange.Start, pastedRange.End) Then
                pathLabel = GetSequenceLabel(fld)
                If Len(pathLabel) > 0 Then
                    Set para = doc.Range(fld.Code.Start, fld.Result.End).Paragraphs(1)
                    ' 粘贴点落在段落中间或段落标记上时，粘贴块首段会与前文合并；
                    ' 把范围夹紧到粘贴块内，书签只覆盖真正粘贴出来的文字。
                    captionStart = para.Range.Start
                    captionEnd = para.Range.End
                    If captionStart < pastedRange.Start Then captionStart = pastedRange.Start
                    If captionEnd > pastedRange.End Then captionEnd = pastedRange.End

                    pathDesc = GetCaptionDescription(doc, captionStart, captionEnd, fld.Result.End)
                    result.Add Array(pathLabel, pathDesc, _
                                     captionStart, captionEnd, _
                                     fld.Result.Start, fld.Result.End)
                End If
            End If
        End If
    Next fld

    Set CollectPastedCaptions = result
End Function

' 题注描述文字 = SEQ 结果末尾到段落文字末尾（不含段落标记），空白折叠
Private Function GetCaptionDescription(ByVal doc As Document, _
                                       ByVal paraStart As Long, ByVal paraEnd As Long, _
                                       ByVal seqResultEnd As Long) As String
    Dim textEnd As Long

    textEnd = paraEnd - 1
    If seqResultEnd <= paraStart Then Exit Function
    If seqResultEnd >= textEnd Then Exit Function

    GetCaptionDescription = NormalizeText(doc.Range(seqResultEnd, textEnd).Text)
End Function

' 匹配粘贴出来的题注：标签与描述都相同才认领，避免把指向别的题注的引用误接过来。
' 粘贴出来的题注是原题注的完整拷贝，描述文字必然一致；匹配不到时宁可不动。
Private Function CaptionMatchKey(ByVal labelName As String, ByVal descText As String) As String
    CaptionMatchKey = labelName & ChrW$(0) & NormalizeText(descText)
End Function

Private Function IndexPastedCaptions(ByVal captions As Collection) As Object
    Dim result As Object
    Dim candidates As Collection
    Dim item As Variant
    Dim key As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For Each item In captions
        key = CaptionMatchKey(CStr(item(0)), CStr(item(1)))
        If Not result.Exists(key) Then
            Set candidates = New Collection
            result.Add key, candidates
        End If
        Set candidates = result(key)
        candidates.Add item
    Next item
    Set IndexPastedCaptions = result
End Function

Private Function IndexSourceCaptionIdentities(ByVal doc As Document, ByVal pastedRange As Range, _
                                               ByVal captionIndex As Object) As Object
    Dim result As Object
    Dim identities As Object
    Dim fld As Field
    Dim para As Paragraph
    Dim labelName As String
    Dim descText As String
    Dim key As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    ' Source uniqueness must include unreferenced captions too: an external REF
    ' must not steal a copied caption merely because its own text is identical.
    For Each fld In doc.Fields
        If fld.Type = wdFieldSequence Then
            If fld.Code.Start < pastedRange.Start Or fld.Code.Start >= pastedRange.End Then
                labelName = GetSequenceLabel(fld)
                Set para = doc.Range(fld.Code.Start, fld.Result.End).Paragraphs(1)
                descText = GetCaptionDescription(doc, para.Range.Start, para.Range.End, fld.Result.End)
                key = CaptionMatchKey(labelName, descText)
                If captionIndex.Exists(key) Then
                    If Not result.Exists(key) Then
                        Set identities = CreateObject("Scripting.Dictionary")
                        result.Add key, identities
                    End If
                    Set identities = result(key)
                    identities(CStr(para.Range.Start)) = True
                End If
            End If
        End If
    Next fld
    Set IndexSourceCaptionIdentities = result
End Function

Private Function MatchPastedCaption(ByVal captionIndex As Object, ByVal used As Object, _
                                    ByVal labelName As String, ByVal descText As String, _
                                    ByRef ambiguousCaptions As Boolean) As Variant
    Dim candidates As Collection
    Dim item As Variant
    Dim key As String
    MatchPastedCaption = Empty
    key = CaptionMatchKey(labelName, descText)
    If Not captionIndex.Exists(key) Then Exit Function
    Set candidates = captionIndex(key)
    ' Check uniqueness before considering claims; a previously used candidate
    ' must not make an originally ambiguous group look unique.
    If candidates.Count <> 1 Then
        ambiguousCaptions = True
        Exit Function
    End If
    item = candidates(1)
    If used.Exists(CStr(item(2))) Then Exit Function
    used.Add CStr(item(2)), True
    MatchPastedCaption = item
End Function


' ==================== 重连引用 ====================

Private Sub RelinkPastedReferences(ByVal doc As Document, ByVal pastedRange As Range, _
                                   ByVal captions As Collection, ByVal fields As Collection, _
                                    ByRef ambiguousCaptions As Boolean)
    Dim fld As Field
    Dim refCount As Long
    Dim oldName As String
    Dim newName As String
    Dim cache As Object
    Dim used As Object
    Dim claimed As Object
    Dim pendingFields As New Collection
    Dim pendingNames As New Collection
    Dim pendingIndex As Long
    Dim names As Variant
    Dim captionIndex As Object
    Dim sourceIdentities As Object

    If fields Is Nothing Then Exit Sub

    For Each fld In fields
        If fld.Type = wdFieldRef Or fld.Type = wdFieldPageRef Then
            refCount = refCount + 1
        End If
    Next fld
    If refCount = 0 Then Exit Sub

    Set cache = CreateObject("Scripting.Dictionary")
    Set used = CreateObject("Scripting.Dictionary")
    Set claimed = CreateObject("Scripting.Dictionary")
    Set captionIndex = IndexPastedCaptions(captions)
    Set sourceIdentities = IndexSourceCaptionIdentities(doc, pastedRange, captionIndex)

    For Each fld In fields
        If fld.Type = wdFieldRef Or fld.Type = wdFieldPageRef Then
            oldName = GetReferencedBookmarkName(fld)
            If Len(oldName) > 0 Then
                If cache.Exists(oldName) Then
                    newName = CStr(cache(oldName))
                Else
                    newName = BuildRelinkedBookmark(doc, pastedRange, captionIndex, sourceIdentities, used, claimed, oldName, ambiguousCaptions)
                    cache.Add oldName, newName
                End If

                If Len(newName) > 0 Then
                    pendingFields.Add fld
                    pendingNames.Add Array(oldName, newName)
                End If
            End If
        End If
    Next fld

    ' Build every bookmark before editing any field code; numeric caption positions
    ' remain valid during planning, and Word moves created bookmarks with edits.
    For pendingIndex = 1 To pendingFields.Count
        Set fld = pendingFields(pendingIndex)
        names = pendingNames(pendingIndex)
        ReplaceBookmarkInCode fld, CStr(names(0)), CStr(names(1))
    Next pendingIndex
End Sub

' 为一条引用找到粘贴出来的对应题注并新建书签；不需要/无法重连时返回空串
Private Function BuildRelinkedBookmark(ByVal doc As Document, ByVal pastedRange As Range, _
                                       ByVal captionIndex As Object, ByVal sourceIdentities As Object, _
                                       ByVal used As Object, ByVal claimed As Object, ByVal oldName As String, _
                                       ByRef ambiguousCaptions As Boolean) As String
    Dim bmRange As Range
    Dim srcPara As Paragraph
    Dim srcLabel As String
    Dim srcDesc As String
    Dim seqStart As Long
    Dim seqEnd As Long
    Dim target As Variant
    Dim newStart As Long
    Dim newEnd As Long
    Dim newName As String

    If Not doc.Bookmarks.Exists(oldName) Then Exit Function

    On Error Resume Next
    Set bmRange = doc.Bookmarks(oldName).Range
    On Error GoTo 0
    If bmRange Is Nothing Then Exit Function

    ' 跨文档粘贴时 Word 会把书签一并带过来，引用已经指向粘贴出来的题注
    If bmRange.Start >= pastedRange.Start And bmRange.Start <= pastedRange.End Then Exit Function

    On Error Resume Next
    Set srcPara = bmRange.Paragraphs(1)
    On Error GoTo 0
    If srcPara Is Nothing Then Exit Function

    srcLabel = FindSequenceInParagraph(srcPara, seqStart, seqEnd)
    If Len(srcLabel) = 0 Then Exit Function

    ' 同一个源题注上的多条引用（如「仅编号」与「整条题注」）应共用同一个目标题注
    If claimed.Exists(CStr(srcPara.Range.Start)) Then
        target = claimed(CStr(srcPara.Range.Start))
    Else
        srcDesc = GetCaptionDescription(doc, srcPara.Range.Start, srcPara.Range.End, seqEnd)
        Dim sourceKey As String
        Dim identities As Object
        sourceKey = CaptionMatchKey(srcLabel, srcDesc)
        If sourceIdentities.Exists(sourceKey) Then
            Set identities = sourceIdentities(sourceKey)
            If identities.Count <> 1 Then
                ambiguousCaptions = True
                Exit Function
            End If
        End If
        target = MatchPastedCaption(captionIndex, used, srcLabel, srcDesc, ambiguousCaptions)
        If IsEmpty(target) Then Exit Function
        claimed.Add CStr(srcPara.Range.Start), target
    End If

    If Not MapCaptionBookmark(bmRange.Start, bmRange.End, _
                              srcPara.Range.Start, srcPara.Range.End, seqStart, seqEnd, _
                              CLng(target(2)), CLng(target(3)), CLng(target(4)), CLng(target(5)), _
                              newStart, newEnd) Then
        Exit Function
    End If

    newName = UniqueBookmarkName(doc)
    On Error Resume Next
    doc.Bookmarks.Add newName, doc.Range(newStart, newEnd)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    BuildRelinkedBookmark = newName
End Function

' 把源书签的形状映射到目标题注上。
' Word 的交叉引用书签只有四种形状，按形状重建比按字符偏移更稳：
' 题注前缀（如带章节号的 STYLEREF）在源与目标之间长度可能不同。
Private Function MapCaptionBookmark(ByVal bmStart As Long, ByVal bmEnd As Long, _
                                    ByVal srcParaStart As Long, ByVal srcParaEnd As Long, _
                                    ByVal srcSeqStart As Long, ByVal srcSeqEnd As Long, _
                                    ByVal tgtParaStart As Long, ByVal tgtParaEnd As Long, _
                                    ByVal tgtSeqStart As Long, ByVal tgtSeqEnd As Long, _
                                    ByRef newStart As Long, ByRef newEnd As Long) As Boolean
    Dim srcTextEnd As Long
    Dim tgtTextEnd As Long

    srcTextEnd = srcParaEnd - 1
    tgtTextEnd = tgtParaEnd - 1

    If bmStart <= srcParaStart And bmEnd >= srcTextEnd Then
        ' 整个题注
        newStart = tgtParaStart
        newEnd = tgtTextEnd
    ElseIf bmStart <= srcParaStart And bmEnd = srcSeqEnd Then
        ' 标签 + 编号
        newStart = tgtParaStart
        newEnd = tgtSeqEnd
    ElseIf bmStart = srcSeqStart And bmEnd = srcSeqEnd Then
        ' 仅编号
        newStart = tgtSeqStart
        newEnd = tgtSeqEnd
    ElseIf bmStart = srcSeqEnd And bmEnd >= srcTextEnd Then
        ' 仅描述文字
        newStart = tgtSeqEnd
        newEnd = tgtTextEnd
    Else
        ' 兜底：以 SEQ 结果末尾为基准做相对偏移
        newStart = tgtSeqEnd + (bmStart - srcSeqEnd)
        newEnd = tgtSeqEnd + (bmEnd - srcSeqEnd)
    End If

    If newStart < tgtParaStart Then newStart = tgtParaStart
    If newEnd > tgtTextEnd Then newEnd = tgtTextEnd
    If newEnd <= newStart Then Exit Function

    MapCaptionBookmark = True
End Function

' 把域代码里的旧书签名替换成新书签名
Private Sub ReplaceBookmarkInCode(ByVal fld As Field, _
                                  ByVal oldName As String, ByVal newName As String)
    Dim codeText As String
    Dim pos As Long

    codeText = fld.Code.Text
    pos = FindWholeToken(codeText, oldName)
    If pos = 0 Then Exit Sub

    On Error Resume Next
    fld.Code.Text = Left$(codeText, pos - 1) & newName & Mid$(codeText, pos + Len(oldName))
    Err.Clear
End Sub

' 在域代码里按「完整片段」查找书签名：前后必须是空白/引号/边界，
' 避免 _Ref123 命中 _Ref1234 这类前缀重叠
Private Function FindWholeToken(ByVal text As String, ByVal token As String) As Long
    Dim pos As Long
    Dim before As String
    Dim after As String

    If Len(token) = 0 Then Exit Function

    pos = InStr(1, text, token, vbTextCompare)
    Do While pos > 0
        If pos = 1 Then
            before = " "
        Else
            before = Mid$(text, pos - 1, 1)
        End If

        If pos + Len(token) > Len(text) Then
            after = " "
        Else
            after = Mid$(text, pos + Len(token), 1)
        End If

        If IsTokenBoundary(before) And IsTokenBoundary(after) Then
            FindWholeToken = pos
            Exit Function
        End If

        pos = InStr(pos + 1, text, token, vbTextCompare)
    Loop
End Function

Private Function IsTokenBoundary(ByVal ch As String) As Boolean
    Select Case AscW(ch)
        Case 7, 9, 10, 13, 32, 34, 160   ' 控制符、空格、双引号、不换行空格
            IsTokenBoundary = True
    End Select
End Function


' ==================== 域代码解析 ====================

' 解析 REF / PAGEREF 域代码里的书签名（兼容带引号的写法）
Private Function GetReferencedBookmarkName(ByVal fld As Field) As String
    Dim codeText As String
    Dim rest As String
    Dim spacePos As Long
    Dim quoteEnd As Long

    codeText = CleanFieldCode(fld.Code.Text)
    spacePos = InStr(1, codeText, " ", vbTextCompare)
    If spacePos = 0 Then Exit Function

    rest = Trim$(Mid$(codeText, spacePos + 1))
    If Len(rest) = 0 Then Exit Function

    If Left$(rest, 1) = """" Then
        quoteEnd = InStr(2, rest, """")
        If quoteEnd > 2 Then GetReferencedBookmarkName = Mid$(rest, 2, quoteEnd - 2)
    Else
        spacePos = InStr(1, rest, " ", vbTextCompare)
        If spacePos > 0 Then rest = Left$(rest, spacePos - 1)
        GetReferencedBookmarkName = rest
    End If
End Function

' 取 SEQ 代码里的标签名：SEQ 标签 \* ARABIC 中的「标签」
Private Function GetSequenceLabel(ByVal fld As Field) As String
    Dim labelName As String

    labelName = GetCodeToken(fld.Code.Text, 1)
    labelName = Replace(labelName, "\", "")
    labelName = Replace(labelName, "*", "")
    GetSequenceLabel = Trim$(labelName)
End Function

' 段落里的第一个 SEQ 域；返回标签名，同时带出结果范围
Private Function FindSequenceInParagraph(ByVal para As Paragraph, _
                                         ByRef seqStart As Long, ByRef seqEnd As Long) As String
    Dim fld As Field
    Dim labelName As String

    seqStart = 0
    seqEnd = 0

    On Error Resume Next
    For Each fld In para.Range.Fields
        If fld.Type = wdFieldSequence Then
            labelName = GetSequenceLabel(fld)
            If Len(labelName) > 0 Then
                seqStart = fld.Result.Start
                seqEnd = fld.Result.End
                FindSequenceInParagraph = labelName
                Exit Function
            End If
        End If
    Next fld
    Err.Clear
End Function

' 按空白切分域代码，返回第 tokenIndex 个非空片段（0 起算）
Private Function GetCodeToken(ByVal codeText As String, ByVal tokenIndex As Long) As String
    Dim parts() As String
    Dim i As Long
    Dim seen As Long

    parts = Split(CleanFieldCode(codeText), " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If seen = tokenIndex Then
                GetCodeToken = parts(i)
                Exit Function
            End If
            seen = seen + 1
        End If
    Next i
End Function


' ==================== 通用小工具 ====================

' 域代码里可能夹带段落标记、域结束符等控制字符，统一折叠为空格
Private Function CleanFieldCode(ByVal codeText As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    result = vbNullString
    For i = 1 To Len(codeText)
        ch = Mid$(codeText, i, 1)
        Select Case AscW(ch)
            Case 7, 9, 10, 13, 32, 160
                result = result & " "
            Case Else
                result = result & ch
        End Select
    Next i
    CleanFieldCode = Trim$(result)
End Function

' 空白折叠并去首尾空格
Private Function NormalizeText(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    Dim lastWasSpace As Boolean

    result = vbNullString
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        Select Case AscW(ch)
            Case 7, 9, 10, 13, 32, 160
                If Not lastWasSpace Then result = result & " "
                lastWasSpace = True
            Case Else
                result = result & ch
                lastWasSpace = False
        End Select
    Next i
    NormalizeText = Trim$(result)
End Function

Private Function IsRangeInside(ByVal startPos As Long, ByVal endPos As Long, _
                               ByVal containerStart As Long, ByVal containerEnd As Long) As Boolean
    IsRangeInside = (startPos >= containerStart And endPos <= containerEnd And endPos > startPos)
End Function

Private Function UniqueBookmarkName(ByVal doc As Document) As String
    Dim candidate As String
    Dim attempts As Long

    Randomize
    For attempts = 1 To 50
        candidate = "_Ref" & Format$(Int(Rnd() * 900000000) + 100000000, "000000000")
        If Not doc.Bookmarks.Exists(candidate) Then
            UniqueBookmarkName = candidate
            Exit Function
        End If
    Next attempts

    ' 兜底：固定前缀 + 递增序号
    attempts = 0
    Do
        attempts = attempts + 1
        candidate = "_RAToolsRef" & CStr(attempts)
    Loop While doc.Bookmarks.Exists(candidate)
    UniqueBookmarkName = candidate
End Function
