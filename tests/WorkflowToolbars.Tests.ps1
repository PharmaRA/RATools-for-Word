param([ValidateSet('Writing','Final','References','Joint')][string]$Kind = 'Writing', [string]$ProgressPath = '', [switch]$CompileOnly)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'scripts/RATools.Build.psm1') -Force
$word = $null
$ownWordProcessId = $null
$ownWordCreation = $null
$hostDoc = $null
$testAddin = $null
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('RATools_WorkflowToolbarTests_' + [guid]::NewGuid().ToString('N'))
function Write-WorkflowTestProgress([string]$Step) {
    [IO.File]::AppendAllText($ProgressPath, $Step + "`r`n", [Text.Encoding]::UTF8)
    Write-Output $Step
}
try {
    [void][IO.Directory]::CreateDirectory($temporaryRoot)
    if (-not $ProgressPath) { $ProgressPath = Join-Path $temporaryRoot 'progress.log' }
    [IO.File]::WriteAllText($ProgressPath, '', [Text.UTF8Encoding]::new($true))
    Write-WorkflowTestProgress 'Start isolated Word instance'
    $wordStartTime = Get-Date
    $word = Start-RAToolsWordSession
    $ownWordCandidates = @(Get-CimInstance -Query "SELECT ProcessId,CreationDate FROM Win32_Process WHERE Name='WINWORD.EXE' AND (CommandLine LIKE '%Automation%' OR CommandLine LIKE '%Embedding%')" | Where-Object { $_.CreationDate -ge $wordStartTime.AddSeconds(-1) })
    if ($ownWordCandidates.Count -ne 1) { throw 'Unable to identify the isolated Word process safely.' }
    $ownWordProcessId = [int]$ownWordCandidates[0].ProcessId
    $ownWordCreation = $ownWordCandidates[0].CreationDate
    $word.Visible = -not $CompileOnly
    $word.AddIns.Unload($false)
    $templatePath = Join-Path $temporaryRoot 'WorkflowToolbarTests.dotm'
    $null = New-RAToolsDotmFromDirectory -SourceDirectory (Join-Path $repoRoot 'dotm') -DestinationPath $templatePath
    $hostDoc = $word.Documents.Open($templatePath, $false, $false, $false)
    $components = $hostDoc.VBProject.VBComponents
    $replaceComponents = @()
    foreach ($component in $components) { if ($component.Type -ne 100) { $replaceComponents += $component.Name } }
    foreach ($componentName in $replaceComponents) { $components.Remove($components.Item($componentName)) }
    Write-WorkflowTestProgress 'Import VBA sources'
    foreach ($source in (Get-RAToolsVbaSourceFiles -RepoRoot $repoRoot)) {
        [IO.File]::AppendAllText($ProgressPath, 'Import ' + $source.Name + "`r`n", [Text.Encoding]::UTF8)
        [void]$hostDoc.VBProject.VBComponents.Import($source.Path)
    }
    Write-WorkflowTestProgress 'Imported VBA sources; install presentation adapters and probe'
    $manager = $hostDoc.VBProject.VBComponents.Item('UI_WorkflowToolbars').CodeModule
    $managerText = $manager.Lines(1,$manager.CountOfLines)
    $managerText = $managerText.Replace('If Len(failureText) > 0 Then MsgBox failureText, vbExclamation, "RATools"', 'WorkflowProbe.lastWorkflowFailure = failureText')
    $managerText = $managerText.Replace('If mActionRunning Then Exit Sub', 'WorkflowProbe.guardWasRunning = mActionRunning: If mActionRunning Then Exit Sub')
    $managerText = $managerText.Replace('ApplyRAToolsStyle Mid$(actionKey, 7)', 'WorkflowProbe.styleDispatchKey = Mid$(actionKey, 7): ApplyRAToolsStyle Mid$(actionKey, 7)')
    if ($Kind -eq 'Final') {
        $managerText = $managerText.Replace('commentCount = doc.Comments.Count', 'If WorkflowProbe.forceSnapshotError Then Err.Raise vbObjectError + 3110, , "snapshot unavailable"' + "`r`n    commentCount = doc.Comments.Count")
        $managerText = $managerText.Replace('Case "save_current": target.Save', 'Case "save_current": WorkflowProbe.SaveCurrentForTest target')
        foreach ($route in @{'detect_clean_versions'='BatchDetectCleanVersions';'detect_highlights'='BatchDetectHighlights';'export_pdf'='BatchConvertWordToPDF';'accept_revisions_comments'='BatchAcceptAndClean';'heading_numbers_text'='ConvertHeadingNumbers'}.GetEnumerator()) {
            $managerText = $managerText.Replace($route.Value, 'WorkflowProbe.RecordFinalAction "' + $route.Key + '"')
        }
    }
    if ($Kind -eq 'References') {
        foreach ($route in @{
            'caption_pair' = 'GenerateCaptionReferencePair Nothing'
            'restore_caption_labels' = 'Wrapper_AddMissingCaptionLabels Nothing'
            'paste_relink' = 'Wrapper_PasteRelinkCrossReference'
            'links_fields_blue' = 'SetHyperlinksAndFieldsToBlue'
            'protect_field_format' = 'ProtectFieldFormat Nothing'
        }.GetEnumerator()) {
            if (-not $managerText.Contains($route.Value)) { throw ('Missing reference route: ' + $route.Key) }
            $managerText = $managerText.Replace($route.Value, 'WorkflowProbe.RecordReferenceAction "' + $route.Key + '"')
        }
    }
    if ($Kind -eq 'Joint') {
        $managerText = $managerText.Replace('MsgBox Err.Description, vbExclamation, "RATools"', 'WorkflowProbe.lastWorkflowFailure = Err.Description')
        foreach ($windowModuleName in @('UI_QuickToolbar','UI_ReviewToolbar')) {
            $windowModule = $hostDoc.VBProject.VBComponents.Item($windowModuleName).CodeModule
            $windowText = [regex]::Replace($windowModule.Lines(1,$windowModule.CountOfLines), '(?m)^\s*MsgBox [^\r\n]+', '    WorkflowProbe.lastWorkflowFailure = Err.Description')
            $windowModule.DeleteLines(1,$windowModule.CountOfLines)
            $windowModule.AddFromString($windowText)
        }
    }
    $manager.DeleteLines(1,$manager.CountOfLines)
    $manager.AddFromString($managerText)
    $tableModule = $hostDoc.VBProject.VBComponents.Item('Macro_BatchAutoFitTables').CodeModule
    $tableText = $tableModule.Lines(1,$tableModule.CountOfLines).Replace('MsgBox ', 'Debug.Print ')
    $tableModule.DeleteLines(1,$tableModule.CountOfLines)
    $tableModule.AddFromString($tableText)
    $probe = $hostDoc.VBProject.VBComponents.Add(1)
    $probe.Name = 'WorkflowProbe'
    $probeSource = @'
Option Explicit
Private checks As Long
Public lastWorkflowFailure As String
Public guardWasRunning As Boolean, styleDispatchKey As String
Public forceSnapshotError As Boolean, cancelSave As Boolean
Public lastFinalAction As String
Private savedLeft As String, savedTop As String
Private savedFinalLeft As String, savedFinalTop As String
Private savedReferenceLeft As String, savedReferenceTop As String
Public lastReferenceAction As String, pasteDispatches As Long
Public lastReferenceDocument As Document
Public forceReferenceReentry As Boolean, forceReferenceError As Boolean
Public Sub RestoreWorkflowTestPositions()
    If Len(savedLeft) = 0 Or Len(savedTop) = 0 Then Exit Sub
    SaveSetting "RATools", "WritingToolbar", "Left", savedLeft
    SaveSetting "RATools", "WritingToolbar", "Top", savedTop
    If Len(savedFinalLeft) > 0 And Len(savedFinalTop) > 0 Then
        SaveSetting "RATools", "FinalToolbar", "Left", savedFinalLeft
        SaveSetting "RATools", "FinalToolbar", "Top", savedFinalTop
    End If
    If Len(savedReferenceLeft) > 0 And Len(savedReferenceTop) > 0 Then
        SaveSetting "RATools", "ReferencesToolbar", "Left", savedReferenceLeft
        SaveSetting "RATools", "ReferencesToolbar", "Top", savedReferenceTop
    End If
End Sub
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    Dim logFile As Integer
    checks = checks + 1
    logFile = FreeFile
    Open "__WORKFLOW_PROGRESS_PATH__" For Append As #logFile
    Print #logFile, CStr(checks) & ": " & message
    Close #logFile
    If Not condition Then Err.Raise vbObjectError + 3100, , message
End Sub
Private Function WorkflowCount() As Long
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmWritingToolbar" Then WorkflowCount = WorkflowCount + 1
    Next item
End Function
Public Function TestWritingToolbar() As String
    Dim doc As Document, other As Document, item As Object, table As Table, secondTable As Table
    Dim beforeText As String, startPos As Long, endPos As Long
    On Error GoTo Failed
    savedLeft = GetSetting("RATools", "WritingToolbar", "Left", CStr(Application.Left + 30))
    savedTop = GetSetting("RATools", "WritingToolbar", "Top", CStr(Application.Top + 80))
    RefreshWorkflowToolbarsIfVisible
    Check WorkflowCount() = 0, "event must not load form"
    Set doc = Documents.Add
    doc.Content.Text = "First paragraph" & vbCr & "Second paragraph" & vbCr
    doc.Styles.Add "正文-F", wdStyleTypeParagraph
    doc.Styles.Add "标题5-F", wdStyleTypeParagraph
    doc.Range(0,5).Select
    startPos = Selection.Start: endPos = Selection.End
    ShowWorkflowToolbar "Writing"
    Check IsUserFormWindowVisible("RATools 撰写工具"), "visible modeless"
    Check GetUserFormWindowOwner("RATools 撰写工具") = 0, "document independent owner"
    Check IsUserFormWindowOnTop("RATools 撰写工具"), "topmost"
    Check Selection.Start = startPos And Selection.End = endPos, "show preserves selection"
    Check frmWritingToolbar.Controls.Count = 31, "24 actions plus document, window menu and five group labels"
    Check InStr(frmWritingToolbar.btn0.ControlTipText, "标题1-F") > 0, "full style name in tooltip"
    For Each item In frmWritingToolbar.Controls
        If TypeName(item) = "CommandButton" Then Check Not item.TakeFocusOnClick, "button retains text selection"
    Next item
    RunWorkflowAction "style:正文-F"
    Check doc.Paragraphs(1).Style = "正文-F", "actual body style; paragraph=" & CStr(doc.Paragraphs(1).Style) & "; range=" & CStr(doc.Range(0,5).Style) & "; resolved=" & ResolveStyleName(doc, "正文-F") & "; guard=" & guardWasRunning & "; key=" & styleDispatchKey & "; targetActive=" & CStr(doc Is ActiveDocument) & "; selection=" & Selection.Start & ":" & Selection.End & "; dispatch=" & lastWorkflowFailure
    RunWorkflowAction "text_blue"
    Check doc.Range(0,5).Font.Color = wdColorBlue, "actual selected text color"
    Check doc.Range(6,10).Font.Color <> wdColorBlue, "outside selection unchanged"
    doc.Paragraphs(1).Format.PageBreakBefore = True
    doc.Paragraphs(2).Format.PageBreakBefore = False
    doc.Content.Select
    RefreshWorkflowToolbarsIfVisible
    Check InStr(frmWritingToolbar.btn19.Caption, "混合") > 0, "mixed pagination state"
    RunWorkflowAction "page_break_before"
    Check Selection.ParagraphFormat.PageBreakBefore = True, "mixed toggles on"
    frmWritingToolbar.cbostyles.ListIndex = 2
    Check doc.Paragraphs(1).Style = "标题5-F", "menu applies real style"
    Check frmWritingToolbar.cbostyles.ListIndex = 0, "menu resets"
    beforeText = doc.Content.Text
    frmWritingToolbar.cbostyles.ListIndex = 1
    Check doc.Content.Text = beforeText, "group heading is inert"
    SetQuickToolbarMsoTestMode True
    frmWritingToolbar.cbotext.ListIndex = 1
    Check GetLastQuickToolbarMsoForTest() = "Superscript", "menu native route"
    frmWritingToolbar.cbolayout.ListIndex = 3
    Check GetLastQuickToolbarMsoForTest() = "SectionBreakInsert", "section menu route"
    SetQuickToolbarMsoTestMode False
    Set table = doc.Tables.Add(doc.Range(doc.Content.End - 1,doc.Content.End - 1),2,2)
    doc.Range(doc.Content.End - 1,doc.Content.End - 1).InsertAfter vbCr
    Set secondTable = doc.Tables.Add(doc.Range(doc.Content.End - 1,doc.Content.End - 1),2,2)
    Check doc.Tables.Count = 2, "two distinct tables for scope check"
    table.AutoFitBehavior wdAutoFitFixed
    secondTable.AutoFitBehavior wdAutoFitFixed
    table.Cell(1,1).Range.Select
    RunWorkflowAction "autofit_table"
    Check table.AllowAutoFit, "real current table autofit"
    Check Not secondTable.AllowAutoFit, "other table remains fixed"
    RunWorkflowAction "autofit_all_tables"
    Check secondTable.AllowAutoFit, "whole document route includes second table"
    ShowQuickToolbar
    ShowReviewToolbar
    Check IsUserFormWindowVisible("RATools"), "quick window coexists"
    Check IsUserFormWindowVisible(REVIEW_TOOLBAR_CAPTION), "review window coexists"
    Set other = Documents.Add
    RefreshWorkflowToolbarsIfVisible
    Check InStr(frmWritingToolbar.lblDocument.Caption,other.Name) > 0, "active document changes"
    other.Close wdDoNotSaveChanges
    doc.Activate
    RunWorkflowAction "invalid-test-action"
    Check frmWritingToolbar.btn0.Enabled, "failure restores controls"
    RunWorkflowAction "text_blue"
    Check frmWritingToolbar.btn0.Enabled, "next action succeeds"
    frmWritingToolbar.Left = 135: frmWritingToolbar.Top = 167
    ToggleWorkflowToolbar "Writing"
    Check Not IsUserFormWindowVisible("RATools 撰写工具"), "toggle hides"
    RefreshWorkflowToolbarsIfVisible
    Check Not IsUserFormWindowVisible("RATools 撰写工具"), "events keep hidden"
    UnloadWorkflowToolbarsOnQuit
    ShowWorkflowToolbar "Writing"
    Check frmWritingToolbar.Left = 135 And frmWritingToolbar.Top = 167, "independent saved position"
    doc.Close wdDoNotSaveChanges
    Check Documents.Count = 0, "real no-document state in loaded add-in"
    RefreshWorkflowToolbarsIfVisible
    Check Not frmWritingToolbar.btn0.Enabled, "no document disables buttons"
    Check Not frmWritingToolbar.cbostyles.Enabled, "no document disables menus"
    Check InStr(frmWritingToolbar.lblDocument.Caption, "没有打开") > 0, "no document label"
    SetQuickToolbarMsoTestMode True
    RunWorkflowAction "mso:Subscript"
    Check GetLastQuickToolbarMsoForTest() <> "Subscript", "no document cannot bypass dispatcher"
    SetQuickToolbarMsoTestMode False
    UnloadWorkflowToolbarsOnQuit
    UnloadReviewToolbarOnQuit
    UnloadQuickToolbarOnQuit
    RestoreWorkflowTestPositions
    TestWritingToolbar = "PASS WritingToolbar (" & checks & " assertions; actual style/color/table, menu native commands recorded)"
    Exit Function
Failed:
    TestWritingToolbar = "FAIL WritingToolbar after " & checks & ": " & Err.Description
End Function
'@
    $finalProbe = @'
Public Sub RecordFinalAction(ByVal actionKey As String)
    lastFinalAction = actionKey
End Sub
Public Sub SaveCurrentForTest(ByVal target As Document)
    If cancelSave Then Err.Raise 4198, , "user cancelled save"
    target.Save
End Sub
Private Function FinalLoadedCount() As Long
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmFinalToolbar" Then FinalLoadedCount = FinalLoadedCount + 1
    Next item
End Function
Public Function TestFinalToolbar() As String
    Dim doc As Document, other As Document, values As Variant
    On Error GoTo Failed
    savedLeft = GetSetting("RATools", "WritingToolbar", "Left", CStr(Application.Left + 30))
    savedTop = GetSetting("RATools", "WritingToolbar", "Top", CStr(Application.Top + 80))
    savedFinalLeft = GetSetting("RATools", "FinalToolbar", "Left", CStr(Application.Left + 30))
    savedFinalTop = GetSetting("RATools", "FinalToolbar", "Top", CStr(Application.Top + 80))
    RefreshWorkflowToolbarsIfVisible
    Check FinalLoadedCount() = 0, "events do not instantiate final form"
    Set doc = Documents.Add
    doc.TrackRevisions = False
    doc.Content.Text = "Final document" & vbCr
    ShowWorkflowToolbar "Final"
    Check frmFinalToolbar.Controls.Count = 21, "12 actions, refresh, window menu and document/status/group labels"
    Check IsUserFormWindowVisible("RATools 定稿工具"), "final visible"
    Check GetUserFormWindowOwner("RATools 定稿工具") = 0, "final owner detached"
    Check InStr(frmFinalToolbar.lblSnapshot.Caption, "修订：0") > 0, "zero revisions is a successful read"
    Check InStr(frmFinalToolbar.lblSnapshot.Caption, "批注：0") > 0, "zero comments is a successful read"
    Check InStr(frmFinalToolbar.lblSnapshotTime.Caption, "读取时间") > 0, "snapshot timestamp"
    doc.TrackRevisions = True
    doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = "Header tracked"
    doc.Comments.Add doc.Range(0,1), "Review comment"
    RunWorkflowAction "refresh_final_state"
    values = ReadWorkflowReviewSnapshot(doc)
    Check CLng(values(0)) > 0, "revision count includes header story"
    Check CLng(values(1)) = 1, "actual comment count"
    Check InStr(frmFinalToolbar.lblSnapshot.Caption, "批注：1") > 0, "panel reports actual count"
    RefreshWorkflowToolbarsIfVisible True
    Check InStr(frmFinalToolbar.lblSnapshotTime.Caption, "待刷新") > 0, "possible edit marks old values stale"
    RunWorkflowAction "refresh_final_state"
    Check InStr(frmFinalToolbar.lblSnapshotTime.Caption, "待刷新") = 0, "explicit refresh clears stale marker"
    forceSnapshotError = True
    frmFinalToolbar.ReadDocumentSnapshot
    Check InStr(frmFinalToolbar.lblSnapshot.Caption, "读取失败") > 0, "failed read is not zero"
    forceSnapshotError = False
    RunWorkflowAction "refresh_final_state"
    Check InStr(frmFinalToolbar.lblSnapshot.Caption, "读取失败") = 0, "read failure can recover"
    ToggleWorkflowToolbar "Final"
    doc.Comments.Add doc.Range(1,2), "Edited while hidden"
    ShowWorkflowToolbar "Final"
    Check InStr(frmFinalToolbar.lblSnapshot.Caption, "批注：2") > 0, "reopen re-reads edits while already unsaved"
    Set other = Documents.Add
    other.TrackRevisions = False
    RefreshWorkflowToolbarsIfVisible
    Check InStr(frmFinalToolbar.lblDocument.Caption, other.Name) > 0, "document switch targets current document"
    Check InStr(frmFinalToolbar.lblSnapshot.Caption, "批注：0") > 0, "document switch re-reads snapshot"
    other.Close wdDoNotSaveChanges
    doc.Activate
    SetQuickToolbarMsoTestMode True
    RunWorkflowAction "mso:TableOfContentsUpdate"
    Check GetLastQuickToolbarMsoForTest() = "TableOfContentsUpdate", "chapter TOC existing native route"
    RunWorkflowAction "mso:TableOfFiguresUpdate"
    Check GetLastQuickToolbarMsoForTest() = "TableOfFiguresUpdate", "figure TOC existing native route"
    RunWorkflowAction "mso:FieldsUpdate"
    Check GetLastQuickToolbarMsoForTest() = "FieldsUpdate", "current field route"
    SetQuickToolbarMsoTestMode False
    cancelSave = True
    doc.Saved = False
    RunWorkflowAction "save_current"
    Check Not doc.Saved, "cancelled save leaves unsaved state"
    Check InStr(frmFinalToolbar.lblDocument.Caption, "未保存") > 0, "cancelled save never reports saved"
    cancelSave = False
    RunWorkflowAction "export_pdf"
    Check lastFinalAction = "export_pdf", "PDF uses existing mode/options entry"
    RunWorkflowAction "detect_clean_versions"
    Check lastFinalAction = "detect_clean_versions", "full clean check remains separate"
    ShowWorkflowToolbar "Writing"
    Check IsUserFormWindowVisible("RATools 撰写工具"), "writing and final coexist"
    frmFinalToolbar.Left = 221: frmFinalToolbar.Top = 181
    ToggleWorkflowToolbar "Final"
    UnloadWorkflowToolbarsOnQuit
    ShowWorkflowToolbar "Final"
    Check frmFinalToolbar.Left = 221 And frmFinalToolbar.Top = 181, "final independent position"
    doc.Close wdDoNotSaveChanges
    RefreshWorkflowToolbarsIfVisible
    Check Documents.Count = 0, "real final no-document state"
    Check Not frmFinalToolbar.btnRefreshState.Enabled And Not frmFinalToolbar.btn0.Enabled, "no-document disables final actions"
    Check InStr(frmFinalToolbar.lblSnapshot.Caption, "—") > 0, "no-document state is unread"
    UnloadWorkflowToolbarsOnQuit
    RestoreWorkflowTestPositions
    TestFinalToolbar = "PASS FinalToolbar (" & checks & " assertions; snapshot/error/stale/save-cancel/native routes)"
    Exit Function
Failed:
    TestFinalToolbar = "FAIL FinalToolbar after " & checks & ": " & Err.Description
End Function
'@
    $referenceProbe = @'
Private Sub SaveReferenceTestPositions()
    savedLeft = GetSetting("RATools", "WritingToolbar", "Left", CStr(Application.Left + 30))
    savedTop = GetSetting("RATools", "WritingToolbar", "Top", CStr(Application.Top + 80))
    savedFinalLeft = GetSetting("RATools", "FinalToolbar", "Left", CStr(Application.Left + 30))
    savedFinalTop = GetSetting("RATools", "FinalToolbar", "Top", CStr(Application.Top + 80))
    savedReferenceLeft = GetSetting("RATools", "ReferencesToolbar", "Left", CStr(Application.Left + 30))
    savedReferenceTop = GetSetting("RATools", "ReferencesToolbar", "Top", CStr(Application.Top + 80))
End Sub
Public Sub RecordReferenceAction(ByVal actionKey As String)
    lastReferenceAction = actionKey
    Set lastReferenceDocument = ActiveDocument
    If actionKey = "paste_relink" Then pasteDispatches = pasteDispatches + 1
    If forceReferenceError Then Err.Raise vbObjectError + 3120, , "reference adapter failure"
    If forceReferenceReentry Then
        forceReferenceReentry = False
        RunWorkflowAction "paste_relink"
    End If
End Sub
Private Sub CheckCaptionPoint(ByVal target As Range, ByVal expectedError As Long, ByVal message As String)
    Dim actualError As Long
    On Error Resume Next
    Err.Clear
    ValidateCaptionPairInsertionPoint target
    actualError = Err.Number
    On Error GoTo 0
    Check actualError = expectedError, message
End Sub
Private Function ReferencesLoadedCount() As Long
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmReferencesToolbar" Then ReferencesLoadedCount = ReferencesLoadedCount + 1
    Next item
End Function
Private Sub ExerciseReferenceMethods(ByVal toolbar As Object)
    Dim original As Document, doc As Document, other As Document, item As Object
    Dim target As Range, table As Table, fld As Field
    SaveReferenceTestPositions
    Set original = ActiveDocument
    Set doc = Documents.Add
    doc.Content.Text = "Reference methods" & vbCr
    toolbar.RefreshDocumentState True
    Check Not toolbar.btn8.Enabled And Not toolbar.cbotable_tools.Enabled, "busy disables reference actions and menu"
    toolbar.RefreshDocumentState False
    Check toolbar.btn8.Enabled And toolbar.cbotable_tools.Enabled, "idle enables reference actions and menu"
    Check toolbar.Controls.Count = 20, "fourteen reference entries plus document, window menu and four group labels"
    Check InStr(toolbar.lblDocument.Caption, doc.Name) > 0, "reference methods follow active document"
    Check InStr(toolbar.btn8.ControlTipText, "无需先粘贴") > 0, "paste tooltip explains one-step operation"
    Check InStr(toolbar.btn4.ControlTipText, "只设置样式") > 0, "caption style is distinct from automatic numbering"
    For Each item In toolbar.Controls
        If TypeName(item) = "CommandButton" Then Check Not item.TakeFocusOnClick, "reference button preserves selection"
    Next item
    CheckCaptionPoint doc.Range(0, 0), 0, "caption accepts collapsed body insertion point"
    CheckCaptionPoint doc.Range(0, 1), vbObjectError + 2401, "caption rejects a text selection"
    Set target = doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Duplicate
    target.Collapse wdCollapseStart
    CheckCaptionPoint target, vbObjectError + 2403, "caption rejects header story"
    Set table = doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 1, 1)
    Set target = table.Cell(1, 1).Range.Duplicate
    target.Collapse wdCollapseStart
    CheckCaptionPoint target, vbObjectError + 2403, "caption rejects table insertion point"
    Set fld = doc.Fields.Add(doc.Range(0, 0), wdFieldDate)
    Set target = doc.Range(fld.Code.Start + 1, fld.Code.Start + 1)
    CheckCaptionPoint target, vbObjectError + 2404, "caption rejects an existing field code"
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).Select
    SetQuickToolbarMsoTestMode True
    toolbar.cbotable_tools.ListIndex = 1
    Check GetLastQuickToolbarMsoForTest() = "", "menu heading is inert"
    toolbar.cbotable_tools.ListIndex = 11
    Check GetLastQuickToolbarMsoForTest() = "TableInsertGallery", "table insert menu retains native route"
    toolbar.cbotable_tools.ListIndex = 13
    Check GetLastQuickToolbarMsoForTest() = "TableRepeatHeaderRows", "repeated header menu retains native route"
    toolbar.cbotable_tools.ListIndex = 17
    Check GetLastQuickToolbarMsoForTest() = "TextWrappingInLineWithText", "image wrapping menu retains native route"
    Check toolbar.cbotable_tools.ListIndex = 0, "reference menu resets after dispatch"
    RunWorkflowAction "mso:CaptionInsert"
    Check GetLastQuickToolbarMsoForTest() = "CaptionInsert", "caption dialog native route"
    RunWorkflowAction "mso:CrossReferenceInsert"
    Check GetLastQuickToolbarMsoForTest() = "CrossReferenceInsert", "cross-reference dialog native route"
    RunWorkflowAction "mso:FieldsUpdate"
    Check GetLastQuickToolbarMsoForTest() = "FieldsUpdate", "current-range field update shares final route"
    RunWorkflowAction "mso:TableOfFiguresInsert"
    Check GetLastQuickToolbarMsoForTest() = "TableOfFiguresInsert", "figure directory insert route"
    SetQuickToolbarMsoTestMode False
    RunWorkflowAction "caption_pair"
    Check lastReferenceAction = "caption_pair", "caption pair uses existing macro entry"
    pasteDispatches = 0
    forceReferenceReentry = True
    RunWorkflowAction "paste_relink"
    Check pasteDispatches = 1 And guardWasRunning, "one paste dispatch and reentry is blocked"
    RunWorkflowAction "restore_caption_labels"
    Check lastReferenceAction = "restore_caption_labels", "caption labels use existing recovery entry"
    Set other = Documents.Add
    toolbar.RefreshDocumentState False
    RunWorkflowAction "links_fields_blue"
    Check lastReferenceDocument Is other, "reference batch route targets newly active document"
    RunWorkflowAction "protect_field_format"
    Check lastReferenceAction = "protect_field_format", "format protection shares existing final entry"
    other.Close wdDoNotSaveChanges
    doc.Activate
    forceReferenceError = True
    RunWorkflowAction "paste_relink"
    Check InStr(lastWorkflowFailure, "reference adapter failure") > 0, "reference failure is captured"
    forceReferenceError = False
    toolbar.RefreshDocumentState False
    RunWorkflowAction "paste_relink"
    Check lastWorkflowFailure = "" And toolbar.btn8.Enabled, "reference dispatcher recovers after failure"
    Set lastReferenceDocument = Nothing
    doc.Close wdDoNotSaveChanges
    original.Activate
    toolbar.RefreshDocumentState False
End Sub
Public Function TestReferencesToolbar() As String
    Dim doc As Document, item As Object, table As Table, otherTable As Table
    Dim startPos As Long, endPos As Long
    On Error GoTo Failed
    SaveReferenceTestPositions
    RefreshWorkflowToolbarsIfVisible
    Check ReferencesLoadedCount() = 0, "events do not instantiate reference form"
    Set doc = Documents.Add
    doc.Content.Text = "Table title" & vbCr & "Figure title" & vbCr
    doc.Styles.Add "表标题-F", wdStyleTypeParagraph
    doc.Styles.Add "图标题-F", wdStyleTypeParagraph
    doc.Range(0, 5).Select
    startPos = Selection.Start: endPos = Selection.End
    ShowWorkflowToolbar "References"
    Check IsUserFormWindowVisible("RATools 图表引用工具"), "references visible modeless"
    Check GetUserFormWindowOwner("RATools 图表引用工具") = 0, "references owner detached"
    Check Selection.Start = startPos And Selection.End = endPos, "reference show preserves selection"
    ExerciseReferenceMethods frmReferencesToolbar
    doc.Range(0, 5).Select
    RunWorkflowAction "style:表标题-F"
    Check doc.Paragraphs(1).Style = "表标题-F", "actual table caption style"
    Check doc.Fields.Count = 0, "caption styling does not create automatic numbering"
    doc.Paragraphs(2).Range.Select
    RunWorkflowAction "style:图标题-F"
    Check doc.Paragraphs(2).Style = "图标题-F", "actual figure caption style"
    Set table = doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 2, 2)
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter vbCr
    Set otherTable = doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 2, 2)
    table.AutoFitBehavior wdAutoFitFixed
    otherTable.AutoFitBehavior wdAutoFitFixed
    table.Cell(1, 1).Range.Select
    RunWorkflowAction "autofit_table"
    Check table.AllowAutoFit And Not otherTable.AllowAutoFit, "references autofit only current table"
    frmReferencesToolbar.Left = 267: frmReferencesToolbar.Top = 193
    ToggleWorkflowToolbar "References"
    Check Not frmReferencesToolbar.Visible, "reference toggle hides"
    RefreshWorkflowToolbarsIfVisible
    Check Not frmReferencesToolbar.Visible, "events keep reference form hidden"
    UnloadWorkflowToolbarsOnQuit
    ShowWorkflowToolbar "References"
    Check frmReferencesToolbar.Left = 267 And frmReferencesToolbar.Top = 193, "references independent saved position"
    doc.Close wdDoNotSaveChanges
    RefreshWorkflowToolbarsIfVisible
    Check Documents.Count = 0, "reference real no-document state"
    Check Not frmReferencesToolbar.btn8.Enabled And Not frmReferencesToolbar.cbotable_tools.Enabled, "no-document disables reference buttons and menu"
    SetQuickToolbarMsoTestMode True
    RunWorkflowAction "mso:CrossReferenceInsert"
    Check GetLastQuickToolbarMsoForTest() = "", "no-document cannot bypass reference dispatcher"
    SetQuickToolbarMsoTestMode False
    UnloadWorkflowToolbarsOnQuit
    RestoreWorkflowTestPositions
    TestReferencesToolbar = "PASS ReferencesToolbar (" & checks & " assertions; real guards/style/table, dialog and paste routes recorded)"
    Exit Function
Failed:
    TestReferencesToolbar = "FAIL ReferencesToolbar after " & checks & ": " & Err.Description
End Function
'@
    $probeSource += [Environment]::NewLine + $referenceProbe
    $jointProbe = @'
Private Function LoadedWindowForTest(ByVal formName As String) As Object
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = formName Then
            Set LoadedWindowForTest = item
            Exit Function
        End If
    Next item
End Function
Private Function WorkflowWindowsForTest() As Collection
    Dim windows As New Collection, item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmWritingToolbar" Or TypeName(item) = "frmFinalToolbar" Or TypeName(item) = "frmReferencesToolbar" Then windows.Add item
    Next item
    Set WorkflowWindowsForTest = windows
End Function
Private Function WorkflowLoadedCountForTest() As Long
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmWritingToolbar" Or TypeName(item) = "frmFinalToolbar" Or TypeName(item) = "frmReferencesToolbar" Then WorkflowLoadedCountForTest = WorkflowLoadedCountForTest + 1
    Next item
End Function
Public Function TestWorkflowJoint() As String
    Dim doc As Document, other As Document, writing As Object, final As Object, references As Object, toolbar As Object, quick As Object, review As Object
    On Error GoTo Failed
    SaveReferenceTestPositions
    RefreshWorkflowToolbarsIfVisible
    Check WorkflowLoadedCountForTest() = 0, "joint events do not instantiate new windows"
    Set doc = Documents.Add
    doc.Content.Text = "Joint toolbar document" & vbCr
    ShowWorkflowToolbar "Writing"
    Set writing = LoadedWindowForTest("frmWritingToolbar")
    Check Not (writing Is Nothing), "writing window loaded through shared entry"
    Check writing.Visible, "writing window shown; " & lastWorkflowFailure
    Check writing.cboToolbars.ListCount = 6, "unified menu has five windows"
    writing.cboToolbars.ListIndex = 4
    Set final = LoadedWindowForTest("frmFinalToolbar")
    Check Not (final Is Nothing), "writing menu opens final window"
    Check final.Visible And writing.Visible, "opening final preserves writing visibility"
    Check writing.cboToolbars.ListIndex = 0, "window menu resets after switch"
    final.cboToolbars.ListIndex = 5
    Set references = LoadedWindowForTest("frmReferencesToolbar")
    Check Not (references Is Nothing), "final menu opens reference window"
    Check references.Visible And final.Visible And writing.Visible, "all three new windows coexist"
    references.cboToolbars.ListIndex = 1
    Set quick = LoadedWindowForTest("frmQuickToolbar")
    Check Not (quick Is Nothing), "unified menu loads current quick instance"
    Check quick.Visible, "unified menu shows current quick instance"
    references.cboToolbars.ListIndex = 3
    Set review = LoadedWindowForTest("frmReviewToolbar")
    Check Not (review Is Nothing), "unified menu loads current review instance"
    Check review.Visible, "unified menu shows current review instance"
    Check WorkflowLoadedCountForTest() = 3, "shared switch keeps one instance per new window"
    Check GetUserFormWindowOwner(writing.Caption) = 0 And GetUserFormWindowOwner(final.Caption) = 0 And GetUserFormWindowOwner(references.Caption) = 0, "new windows detach from document owner"
    For Each toolbar In WorkflowWindowsForTest()
        Check toolbar.cboToolbars.ListCount = 6 And toolbar.cboToolbars.ListIndex = 0, "same initialized menu on each new window"
        toolbar.RefreshDocumentState True
        Check Not toolbar.cboToolbars.Enabled, "busy disables window menu"
        toolbar.RefreshDocumentState False
    Next toolbar
    Set other = Documents.Add
    RefreshWorkflowToolbarsIfVisible
    For Each toolbar In WorkflowWindowsForTest()
        Check InStr(toolbar.lblDocument.Caption, other.Name) > 0, "joint windows follow current document"
    Next toolbar
    other.Close wdDoNotSaveChanges
    doc.Activate
    writing.Left = 135: writing.Top = 167
    final.Left = 221: final.Top = 181
    references.Left = 267: references.Top = 193
    writing.cboToolbars.ListIndex = 2
    Check Not writing.Visible And final.Visible And references.Visible, "self toggle hides only selected window"
    RefreshWorkflowToolbarsIfVisible
    Check Not writing.Visible, "joint event keeps selected window hidden"
    final.cboToolbars.ListIndex = 2
    Check writing.Visible And writing.Left = 135 And writing.Top = 167, "other window reopens writing in saved position"
    UnloadWorkflowToolbarsOnQuit
    Check WorkflowLoadedCountForTest() = 0, "joint unload removes all new window instances"
    Check quick.Visible And review.Visible, "new window unload preserves current legacy instances"
    ShowWorkflowToolbar "Writing"
    ShowWorkflowToolbar "Final"
    ShowWorkflowToolbar "References"
    Set writing = LoadedWindowForTest("frmWritingToolbar")
    Set final = LoadedWindowForTest("frmFinalToolbar")
    Set references = LoadedWindowForTest("frmReferencesToolbar")
    Check writing.Left = 135 And writing.Top = 167, "writing independent position survives reload"
    Check final.Left = 221 And final.Top = 181, "final independent position survives reload"
    Check references.Left = 267 And references.Top = 193, "reference independent position survives reload"
    doc.Close wdDoNotSaveChanges
    RefreshWorkflowToolbarsIfVisible
    Check Documents.Count = 0, "joint add-in remains available with no document"
    For Each toolbar In WorkflowWindowsForTest()
        Check toolbar.cboToolbars.Enabled, "window switch is available without a document"
        Check Not toolbar.Controls("btn0").Enabled, "document action is disabled without a document"
    Next toolbar
    references.cboToolbars.ListIndex = 2
    Check Not writing.Visible, "no-document menu hides another window"
    references.cboToolbars.ListIndex = 2
    Check writing.Visible, "no-document menu reopens another window"
    SetQuickToolbarMsoTestMode True
    RunWorkflowAction "mso:CaptionInsert"
    Check GetLastQuickToolbarMsoForTest() = "", "no-document switch does not bypass action guard"
    SetQuickToolbarMsoTestMode False
    UnloadWorkflowToolbarsOnQuit
    UnloadReviewToolbarOnQuit
    UnloadQuickToolbarOnQuit
    Check WorkflowLoadedCountForTest() = 0, "joint final teardown unloads new windows"
    Set quick = LoadedWindowForTest("frmQuickToolbar")
    Set review = LoadedWindowForTest("frmReviewToolbar")
    Check quick Is Nothing, "joint teardown unloads current quick instance"
    Check review Is Nothing, "joint teardown unloads current review instance"
    RestoreWorkflowTestPositions
    TestWorkflowJoint = "PASS WorkflowJoint (" & checks & " assertions; actual five-window menus, coexistence, positions and lifecycle)"
    Exit Function
Failed:
    TestWorkflowJoint = "FAIL WorkflowJoint after " & checks & ": " & Err.Description
End Function
'@
    $probeSource += [Environment]::NewLine + $jointProbe
    $methodProbe = @'
Public Function TestWorkflowMethods() As String
    Dim toolbar As Object
    On Error GoTo Failed
    savedLeft = GetSetting("RATools", "WritingToolbar", "Left", CStr(Application.Left + 30))
    savedTop = GetSetting("RATools", "WritingToolbar", "Top", CStr(Application.Top + 80))
    savedFinalLeft = GetSetting("RATools", "FinalToolbar", "Left", CStr(Application.Left + 30))
    savedFinalTop = GetSetting("RATools", "FinalToolbar", "Top", CStr(Application.Top + 80))
    Set toolbar = VBA.UserForms.Add("frm__WORKFLOW_KIND__Toolbar")
    If "__WORKFLOW_KIND__" = "References" Then ExerciseReferenceMethods toolbar
    toolbar.RefreshDocumentState True
    If "__WORKFLOW_KIND__" = "Final" Then
        toolbar.ReadDocumentSnapshot
        toolbar.RefreshDocumentState False, True
        Check InStr(toolbar.lblSnapshotTime.Caption, "待刷新") > 0, "executed final snapshot and stale methods"
    End If
    Check InStr(toolbar.lblDocument.Caption, ActiveDocument.Name) > 0, "executed document state method"
    UnloadWorkflowToolbarsOnQuit
    RestoreWorkflowTestPositions
    TestWorkflowMethods = "PASS WorkflowMethods (__WORKFLOW_KIND__; " & checks & " assertions; direct state and route execution)"
    Exit Function
Failed:
    TestWorkflowMethods = "FAIL WorkflowMethods (__WORKFLOW_KIND__): " & Err.Description
End Function
'@
    $finalProbe += "`r`n" + $methodProbe.Replace('__WORKFLOW_KIND__', $Kind)
    $probeSource += "`r`n" + $finalProbe
    $probe.CodeModule.AddFromString($probeSource.Replace('__WORKFLOW_PROGRESS_PATH__', $ProgressPath.Replace('"', '""')))
    if ($CompileOnly -and $Kind -ne 'Joint') {
        Write-WorkflowTestProgress ('Execute ' + $Kind + ' state methods in hidden Word')
        $result = [string]$word.Run('TestWorkflowMethods')
    } else {
    Write-WorkflowTestProgress 'Save test add-in'
    $hostDoc.Save()
    $hostDoc.Saved = $true
    Write-WorkflowTestProgress 'Close test host document'
    $hostDoc.Close(0)
    $hostDoc = $null
    Write-WorkflowTestProgress 'Load test add-in'
    $testAddin = $word.AddIns.Add($templatePath, $true)
    $testAddin.Installed = $true
    $testMethod = switch ($Kind) { 'Final' { 'TestFinalToolbar' }; 'References' { 'TestReferencesToolbar' }; 'Joint' { 'TestWorkflowJoint' }; default { 'TestWritingToolbar' } }
    Write-WorkflowTestProgress ('Run ' + $Kind + ' toolbar assertions')
    $result = [string]$word.Run($testMethod)
    }
    Write-Output $result
    if (-not $result.StartsWith('PASS')) { throw $result }
}
finally {
    try {
        if ($null -ne $word) {
            foreach ($cleanupMacro in @('UnloadWorkflowToolbarsOnQuit','UnloadReviewToolbarOnQuit','UnloadQuickToolbarOnQuit','RestoreWorkflowTestPositions')) {
                try { $null = $word.Run($cleanupMacro) } catch { }
            }
            foreach ($document in @($word.Documents)) {
                try { $document.Close(0) } catch { Write-Warning ('Test document close failed: ' + $_.Exception.Message) }
            }
            if ($null -ne $testAddin) {
                try { $testAddin.Installed = $false } catch { Write-Warning ('Test add-in unload failed: ' + $_.Exception.Message) }
            }
        }
    }
    finally {
        try {
            try { Stop-RAToolsWordSession -Word $word }
            finally {
                if ($null -ne $ownWordProcessId) {
                    $remainingWord = Get-Process -Id $ownWordProcessId -ErrorAction SilentlyContinue
                    if ($null -ne $remainingWord) {
                        if ($remainingWord.ProcessName -ne 'WINWORD' -or [math]::Abs(($remainingWord.StartTime - $ownWordCreation).TotalSeconds) -gt 2) { throw 'Word process identity changed; refusing forced cleanup.' }
                        try {
                            $remainingWord.Kill()
                            $remainingWord.WaitForExit(10000) | Out-Null
                        } catch {
                            if (-not $remainingWord.HasExited) { throw }
                        }
                    }
                }
            }
        }
        finally { Move-RAToolsToRecycleBin -Path $temporaryRoot -SafeRoot ([IO.Path]::GetTempPath()) }
    }
}
