#requires -version 5.1
param([string]$PackagePath)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../scripts/RATools.Build.psm1') -Force
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$modulePath = Join-Path $repoRoot 'modules/Macro_CaptionReferencePair.bas'
$formPath = Join-Path $repoRoot 'userforms/frmCaptionReference.frm'
if (-not (Test-Path -LiteralPath $modulePath)) {
    throw 'Caption/reference generation is not implemented: no callable feature module.'
}
$word = $null
$hostDoc = $null
try {
    $word = Start-RAToolsWordSession
    if ($PackagePath) {
        $package = (Resolve-Path -LiteralPath $PackagePath).Path
        $hostDoc = $word.Documents.Open($package, $false, $true, $false)
    } else {
        $hostDoc = $word.Documents.Add()
        [void]$hostDoc.VBProject.VBComponents.Import($modulePath)
        [void]$hostDoc.VBProject.VBComponents.Import($formPath)
        [void]$hostDoc.VBProject.VBComponents.Import((Join-Path $repoRoot 'modules/Macro_PasteRelinkCrossReference.bas'))
    }
    $probe = $hostDoc.VBProject.VBComponents.Add(1)
    $probe.Name = 'CaptionPairBehaviorTests'
    $probe.CodeModule.AddFromString(@'
Option Explicit

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2000, "CaptionPairTests", message
End Sub

Private Function FieldOfType(ByVal target As Range, ByVal kind As Long) As Field
    Dim item As Field
    For Each item In target.Fields
        If item.Type = kind Then
            Set FieldOfType = item
            Exit Function
        End If
    Next
End Function

Private Sub RefreshPairs(ByVal doc As Document)
    Dim item As Field
    For Each item In doc.Fields
        If item.Type = wdFieldSequence Then item.Update
    Next
    For Each item In doc.Fields
        If item.Type = wdFieldRef Then item.Update
    Next
End Sub

Public Function TestCaptionPairGeneration() As String
    Dim doc As Document, pair As Range, seq As Field, ref As Field
    Dim beforeText As String, countBefore As Long, code As String, name As String
    On Error GoTo Failed
    Set doc = Documents.Add
    doc.Content.Text = "Before" & vbCr & "After" & vbCr
    beforeText = doc.Content.Text
    Set pair = InsertCaptionReferencePair(doc.Range(7, 7), "Table", "Stability results")
    Set seq = FieldOfType(pair, wdFieldSequence)
    Set ref = FieldOfType(pair, wdFieldRef)
    Check Not seq Is Nothing, "Caption must contain a real SEQ field"
    Check Not ref Is Nothing, "Reference must contain a real REF field"
    Check ref.Result.Text = "Table 1", "Reference must exclude the caption title"
    Check pair.Text = "Table 1" & vbCr & "Table 1 Stability results" & vbCr, "Reference must precede caption in adjacent paragraphs"
    Check ref.Result.Start < seq.Result.Start, "REF field must be before the caption SEQ field"
    Check Left$(doc.Content.Text, 7) = "Before" & vbCr, "Preceding body must be preserved"
    Check InStr(doc.Content.Text, "After" & vbCr) > 0, "Following body must be preserved"
    Check InStr(ref.Code.Text, "\h") > 0, "Reference must be navigable"
    Check UBound(doc.GetCrossReferenceItems("Table")) = 1, "Caption must appear in Word's native reference list"
    Check doc.Undo, "Generation must create an undo record"
    Check doc.Content.Text = beforeText, "One undo must restore original content"
    Check doc.Fields.Count = 0, "One undo must remove both fields"
    doc.Close wdDoNotSaveChanges
    TestCaptionPairGeneration = "PASS adjacent pair, native list, hyperlink, preservation, single undo"
    Exit Function
Failed:
    TestCaptionPairGeneration = "FAIL " & Err.Description
    On Error Resume Next
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairFormatting() As String
    Dim doc As Document, pair As Range, seq As Field, ref As Field, number As Variant
    Dim sequenceCode As String
    On Error GoTo Failed
    Set doc = Documents.Add
    doc.Styles(wdStyleNormal).Font.Color = RGB(0, 0, 0)
    doc.Styles(wdStyleCaption).Font.Color = RGB(0, 0, 0)
    doc.Content.Text = "Before" & vbCr & "After" & vbCr
    Set pair = InsertCaptionReferencePair(doc.Range(7, 7), "Table", "Formatting stays separate")
    Set seq = FieldOfType(pair, wdFieldSequence)
    Set ref = FieldOfType(pair, wdFieldRef)
    Check ref.Result.Font.Color = RGB(0, 0, 255), "Generated reference must default to standard blue"
    Check Not ref.Locked, "Formatting protection must allow reference values to update"
    Check InStr(1, ref.Code.Text, "\* MERGEFORMAT", vbTextCompare) > 0, "Generated reference must opt into Word field format preservation"
    Check seq.Result.Font.Color = RGB(0, 0, 0), "Caption must not inherit the blue reference formatting"
    Check doc.Paragraphs(1).Range.Font.Color = RGB(0, 0, 0), "Existing text must keep its color"

    ref.Result.Font.Bold = True
    ref.Result.Font.Italic = True
    ref.Result.Font.Size = 14
    seq.Result.Font.Color = RGB(255, 0, 0)
    seq.Result.Font.Bold = False
    seq.Result.Font.Italic = False
    seq.Result.Font.Size = 10
    sequenceCode = seq.Code.Text
    For Each number In Array(9, 10, 100)
        seq.Code.Text = sequenceCode & " \r " & CStr(number)
        seq.Update
        ref.Update
        Check ref.Result.Text = "Table " & CStr(number), "Protected reference must update to the new caption number"
        Check ref.Result.Font.Color = RGB(0, 0, 255), "Blue must survive repeated updates and number-width changes"
        Check ref.Result.Font.Bold = True And ref.Result.Font.Italic = True And ref.Result.Font.Size = 14, "User font formatting must survive field updates"
    Next number
    doc.Close wdDoNotSaveChanges
    TestCaptionPairFormatting = "PASS standard blue, format preservation, live renumbering and unchanged caption/body colors"
    Exit Function
Failed:
    TestCaptionPairFormatting = "FAIL " & Err.Description
    On Error Resume Next
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairSequencesAndMoves() As String
    Dim doc As Document, first As Range, second As Range, cn As Range, custom As Range
    Dim firstRef As Field, secondRef As Field, moving As Range, dest As Range, refs As Collection
    Dim f As Field, labelsBefore As Long, cnLabel As String, stepName As String
    On Error GoTo Failed
    stepName = "create pairs"
    Set doc = Documents.Add
    doc.Content.Text = "Introduction" & vbCr
    Set first = InsertCaptionReferencePair(doc.Range(doc.Content.End - 1, doc.Content.End - 1), "Table", "Same title")
    Set firstRef = FieldOfType(first, wdFieldRef)
    Set second = InsertCaptionReferencePair(doc.Range(doc.Content.End - 1, doc.Content.End - 1), "Table", "Same title")
    Set secondRef = FieldOfType(second, wdFieldRef)
    Check firstRef.Result.Text = "Table 1", "First pair numbering"
    Check secondRef.Result.Text = "Table 2", "Same label must continue numbering"
    Check firstRef.Code.Text <> secondRef.Code.Text, "Duplicate titles must have independent targets"
    cnLabel = ChrW(34920)
    Set cn = InsertCaptionReferencePair(doc.Range(doc.Content.End - 1, doc.Content.End - 1), cnLabel, "Chinese label")
    Check FieldOfType(cn, wdFieldRef).Result.Text = cnLabel & " 1", "Chinese and English sequences must be independent"
    Set custom = InsertCaptionReferencePair(doc.Range(doc.Content.End - 1, doc.Content.End - 1), "Appendix Table", "Custom title")
    Check FieldOfType(custom, wdFieldRef).Result.Text = "Appendix Table 1", "Custom labels with spaces must work"
    labelsBefore = Application.CaptionLabels.Count
    Set custom = InsertCaptionReferencePair(doc.Range(doc.Content.End - 1, doc.Content.End - 1), "Appendix Table", "")
    Check Application.CaptionLabels.Count = labelsBefore, "Existing label must be reused"
    Check FieldOfType(custom, wdFieldRef).Result.Text = "Appendix Table 2", "Empty title must support a fill-in-later pair"

    Set moving = FieldOfType(first, wdFieldSequence).Result.Paragraphs(1).Range.Duplicate
    stepName = "cut caption"
    moving.Cut
    Set dest = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    stepName = "paste caption"
    dest.Paste
    stepName = "refresh moved caption"
    RefreshPairs doc
    Set refs = New Collection
    For Each f In doc.Fields
        If f.Type = wdFieldRef Then refs.Add f
    Next
    Check refs(1).Result.Text = "Table 2", "Moving first caption must preserve its reference target; actual=[" & refs(1).Result.Text & "]; document=[" & Replace(doc.Content.Text, vbCr, "|") & "]"
    Check refs(2).Result.Text = "Table 1", "Other pair must follow its own caption after reordering"
    Set moving = refs(1).Result.Paragraphs(1).Range.Duplicate
    stepName = "cut reference"
    moving.Cut
    Set dest = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    stepName = "paste reference"
    dest.Paste
    stepName = "refresh moved reference"
    RefreshPairs doc
    Set f = doc.Fields(doc.Fields.Count)
    Check f.Type = wdFieldRef And f.Result.Text = "Table 2", "Moving the reference must preserve its target"
    doc.Close wdDoNotSaveChanges
    TestCaptionPairSequencesAndMoves = "PASS numbering, duplicate titles, Chinese/custom labels, caption and reference cut/paste"
    Exit Function
Failed:
    TestCaptionPairSequencesAndMoves = "FAIL [" & stepName & "] " & Err.Description
    On Error Resume Next
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairInputSafety() As String
    Dim doc As Document, pair As Range, original As String, failed As Boolean, tbl As Table
    On Error GoTo Failed
    Set doc = Documents.Add
    doc.Content.Text = "LeftRight" & vbCr
    Set pair = InsertCaptionReferencePair(doc.Range(4, 4), "Figure", "Mid paragraph")
    Check InStr(doc.Content.Text, "Left" & vbCr & "Figure 1" & vbCr & "Figure 1 Mid paragraph" & vbCr & "Right") = 1, "Inserting mid-paragraph must preserve both halves with the reference first"
    doc.Undo
    original = doc.Content.Text
    On Error Resume Next
    Set pair = InsertCaptionReferencePair(doc.Range(0, 4), "Table", "Must not replace selection")
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo Failed
    Check failed And doc.Content.Text = original, "Noncollapsed selection must be refused without replacement"
    On Error Resume Next
    Set pair = InsertCaptionReferencePair(doc.Range(0, 0), " ", "Invalid label")
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo Failed
    Check failed And doc.Content.Text = original, "Blank label must not change document"
    On Error Resume Next
    Set pair = InsertCaptionReferencePair(doc.Range(0, 0), "Table", "Bad" & vbCr & "title")
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo Failed
    Check failed And doc.Content.Text = original, "Multiline title must not damage paragraph boundaries"
    doc.Protect wdAllowOnlyReading, NoReset:=True
    On Error Resume Next
    Set pair = InsertCaptionReferencePair(doc.Range(0, 0), "Table", "Protected")
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo Failed
    Check failed And doc.Content.Text = original, "Protected document must remain unchanged"
    doc.Close wdDoNotSaveChanges
    TestCaptionPairInputSafety = "PASS mid-paragraph, selection, invalid input and protection safeguards"
    Exit Function
Failed:
    TestCaptionPairInputSafety = "FAIL " & Err.Description
    On Error Resume Next
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairFormCancel() As String
    Dim doc As Document, ui As frmCaptionReference, original As String, labelsBefore As Long
    Dim i As Long, hasTable As Boolean
    On Error GoTo Failed
    Set doc = Documents.Add
    doc.Content.Text = "Unchanged" & vbCr
    original = doc.Content.Text
    labelsBefore = CaptionLabels.Count
    Set ui = New frmCaptionReference
    ui.SetInsertionPoint doc.Range(0, 0)
    Check ui.Controls("cboLabel").ListCount = CaptionLabels.Count, "Form must offer native labels"
    ui.AddLabelChoice "DraftLabelForCancel"
    Check ui.Controls("cboLabel").Value = "DraftLabelForCancel", "New label must be selectable"
    ui.Controls("txtTitle").Text = "Cancelled title"
    ui.Controls("btnCancel").Value = True
    Check doc.Content.Text = original And CaptionLabels.Count = labelsBefore, "Cancelling must not insert or register draft labels"
    Unload ui
    doc.Close wdDoNotSaveChanges
    TestCaptionPairFormCancel = "PASS form native labels, draft label and cancel without side effects"
    Exit Function
Failed:
    TestCaptionPairFormCancel = "FAIL " & Err.Description
    On Error Resume Next
    Unload ui
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairFormGenerate() As String
    Dim doc As Document, ui As frmCaptionReference, previousLabel As String, ref As Field
    On Error GoTo Failed
    previousLabel = GetSetting("RATools", "CaptionReferencePair", "LastLabel", "")
    Set doc = Documents.Add
    Set ui = New frmCaptionReference
    ui.SetInsertionPoint doc.Range(0, 0)
    ui.AddLabelChoice "FormCaptionPair"
    ui.Controls("txtTitle").Text = "Created through the button"
    ui.Controls("btnGenerate").Value = True
    Set ref = FieldOfType(doc.Content, wdFieldRef)
    Check Not ref Is Nothing, "Generate button must create the pair"
    Check ref.Result.Text = "FormCaptionPair 1", "Generate button must use selected custom label"
    Unload ui
    Set ui = New frmCaptionReference
    Check ui.Controls("cboLabel").Value = "FormCaptionPair", "Reopened form must remember the successful label"
    Unload ui
    SaveSetting "RATools", "CaptionReferencePair", "LastLabel", previousLabel
    doc.Close wdDoNotSaveChanges
    TestCaptionPairFormGenerate = "PASS form generate event, custom label and remembered selection"
    Exit Function
Failed:
    TestCaptionPairFormGenerate = "FAIL " & Err.Description
    On Error Resume Next
    Unload ui
    SaveSetting "RATools", "CaptionReferencePair", "LastLabel", previousLabel
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairUnsupportedPositions() As String
    Dim doc As Document, pair As Range, tbl As Table, f As Field
    Dim original As String, failed As Boolean, point As Range
    On Error GoTo Failed
    Set doc = Documents.Add
    Set tbl = doc.Tables.Add(doc.Range(0, 0), 1, 1)
    original = doc.Content.Text
    On Error Resume Next
    Set pair = InsertCaptionReferencePair(tbl.Cell(1, 1).Range.Duplicate, "Table", "Inside table")
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo Failed
    Check failed And doc.Content.Text = original, "Table selection must not be replaced"
    Set point = tbl.Cell(1, 1).Range.Duplicate
    point.Collapse wdCollapseStart
    On Error Resume Next
    Set pair = InsertCaptionReferencePair(point, "Table", "Inside table")
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo Failed
    Check failed And doc.Content.Text = original, "Table cursor must be refused"
    Set point = doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Duplicate
    point.Collapse wdCollapseStart
    On Error Resume Next
    Set pair = InsertCaptionReferencePair(point, "Table", "Header")
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo Failed
    Check failed And doc.Content.Text = original, "Header cursor must be refused"
    Set point = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set f = doc.Fields.Add(point, wdFieldQuote, """Existing field text""", False)
    original = doc.Content.Text
    Set point = doc.Range(f.Result.Start + 2, f.Result.Start + 2)
    On Error Resume Next
    Set pair = InsertCaptionReferencePair(point, "Table", "Inside field")
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo Failed
    Check failed And doc.Content.Text = original, "Existing field must not be split"
    doc.Close wdDoNotSaveChanges
    TestCaptionPairUnsupportedPositions = "PASS table, header and existing-field guards"
    Exit Function
Failed:
    TestCaptionPairUnsupportedPositions = "FAIL " & Err.Description
    On Error Resume Next
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairPasteCompatibility() As String
    Dim doc As Document, pair As Range, originalRef As Field, latestRef As Field, f As Field
    Dim refs As Long
    On Error GoTo Failed
    Set doc = Documents.Add
    Set pair = InsertCaptionReferencePair(doc.Range(0, 0), "Table", "Copy this pair")
    Set originalRef = FieldOfType(pair, wdFieldRef)
    pair.Copy
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).Select
    EnablePasteRelinkQuietMode
    Core_PasteRelinkCrossReference
    Check Len(GetLastPasteRelinkErrorForTest()) = 0, "Paste and relink must accept generated pairs"
    For Each f In doc.Fields
        If f.Type = wdFieldRef Then
            refs = refs + 1
            Set latestRef = f
        End If
    Next
    Check refs = 2, "Pasting a generated pair must produce two references"
    Check originalRef.Result.Text = "Table 1", "Paste must preserve the original pair"
    Check latestRef.Result.Text = "Table 2", "Paste must point to the new caption"
    Check originalRef.Code.Text <> latestRef.Code.Text, "Paste must create a separate target"
    doc.Close wdDoNotSaveChanges
    TestCaptionPairPasteCompatibility = "PASS generated pair copied through existing paste-and-relink"
    Exit Function
Failed:
    TestCaptionPairPasteCompatibility = "FAIL " & Err.Description
    On Error Resume Next
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairNativeNumbering() As String
    Dim doc As Document, pair As Range, f As Field, label As CaptionLabel, i As Long
    On Error GoTo Failed
    Set doc = Documents.Add
    For i = 1 To 12
        Set pair = InsertCaptionReferencePair(doc.Range(doc.Content.End - 1, doc.Content.End - 1), "Table", "Repeated title")
    Next i
    Check FieldOfType(pair, wdFieldRef).Result.Text = "Table 12", "Double digit sequence must preserve bookmark boundaries"
    Set label = CaptionLabels.Add("RomanPairTest")
    label.NumberStyle = wdCaptionNumberStyleUppercaseRoman
    Set pair = InsertCaptionReferencePair(doc.Range(doc.Content.End - 1, doc.Content.End - 1), label.Name, "Roman caption")
    Check FieldOfType(pair, wdFieldRef).Result.Text = "RomanPairTest I", "Native label numbering settings must be respected"
    doc.Close wdDoNotSaveChanges
    TestCaptionPairNativeNumbering = "PASS double digit sequences and native Roman numbering"
    Exit Function
Failed:
    TestCaptionPairNativeNumbering = "FAIL " & Err.Description
    On Error Resume Next
    doc.Close wdDoNotSaveChanges
End Function

Public Function TestCaptionPairOrphanIsolation() As String
    Dim doc As Document, pair As Range, oldRef As Field, oldName As String, newRef As Field
    On Error GoTo Failed
    Set doc = Documents.Add
    Set pair = InsertCaptionReferencePair(doc.Range(0, 0), "Table", "Old caption")
    Set oldRef = FieldOfType(pair, wdFieldRef)
    oldName = Split(Trim$(oldRef.Code.Text), " ")(1)
    FieldOfType(pair, wdFieldSequence).Result.Paragraphs(1).Range.Text = vbNullString
    oldRef.Update
    Check Not doc.Bookmarks.Exists(oldName), "Fixture must retain a REF whose caption no longer exists"
    Set pair = InsertCaptionReferencePair(doc.Range(doc.Content.End - 1, doc.Content.End - 1), "Figure", "Unrelated new caption")
    Set newRef = FieldOfType(pair, wdFieldRef)
    oldRef.Update
    Check Not doc.Bookmarks.Exists(oldName), "New captions must not reuse orphaned reference targets"
    Check oldRef.Result.Text <> "Figure 1", "Broken old reference must not silently resolve to the new caption"
    Check newRef.Result.Text = "Figure 1", "New reference must resolve to its own caption"
    doc.Close wdDoNotSaveChanges
    TestCaptionPairOrphanIsolation = "PASS orphaned references remain isolated from new captions"
    Exit Function
Failed:
    TestCaptionPairOrphanIsolation = "FAIL " & Err.Description
    On Error Resume Next
    doc.Close wdDoNotSaveChanges
End Function
'@)
    foreach ($test in @('TestCaptionPairGeneration', 'TestCaptionPairFormatting', 'TestCaptionPairSequencesAndMoves', 'TestCaptionPairInputSafety', 'TestCaptionPairFormCancel', 'TestCaptionPairFormGenerate', 'TestCaptionPairUnsupportedPositions', 'TestCaptionPairPasteCompatibility', 'TestCaptionPairNativeNumbering', 'TestCaptionPairOrphanIsolation')) {
        $result = $word.Run($test)
        Write-Output "$test`: $result"
        if (-not $result.StartsWith('PASS ')) { throw "$test`: $result" }
    }
} finally {
    if ($hostDoc) { $hostDoc.Close(0) }
    Stop-RAToolsWordSession $word
}
Write-Output 'PASS CaptionReferencePair.Tests'
