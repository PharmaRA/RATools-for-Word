param([ValidateSet('Writing')][string]$Kind = 'Writing', [string]$ProgressPath = '')
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
    $word.Visible = $true
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
Private savedLeft As String, savedTop As String
Public Sub RestoreWorkflowTestPositions()
    If Len(savedLeft) = 0 Or Len(savedTop) = 0 Then Exit Sub
    SaveSetting "RATools", "WritingToolbar", "Left", savedLeft
    SaveSetting "RATools", "WritingToolbar", "Top", savedTop
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
    Check frmWritingToolbar.Controls.Count = 30, "24 actions plus document and five group labels"
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
    $probe.CodeModule.AddFromString($probeSource.Replace('__WORKFLOW_PROGRESS_PATH__', $ProgressPath.Replace('"', '""')))
    Write-WorkflowTestProgress 'Save test add-in'
    $hostDoc.Save()
    $hostDoc.Saved = $true
    Write-WorkflowTestProgress 'Close test host document'
    $hostDoc.Close(0)
    $hostDoc = $null
    Write-WorkflowTestProgress 'Load test add-in'
    $testAddin = $word.AddIns.Add($templatePath, $true)
    $testAddin.Installed = $true
    Write-WorkflowTestProgress 'Run writing toolbar assertions'
    $result = [string]$word.Run('TestWritingToolbar')
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
                        Stop-Process -Id $ownWordProcessId -Force -ErrorAction Stop
                        $remainingWord.WaitForExit(10000) | Out-Null
                    }
                }
            }
        }
        finally { Move-RAToolsToRecycleBin -Path $temporaryRoot -SafeRoot ([IO.Path]::GetTempPath()) }
    }
}
