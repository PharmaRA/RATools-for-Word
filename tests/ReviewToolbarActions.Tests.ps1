$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'scripts/RATools.Build.psm1') -Force
$word = $null
try {
    $word = Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $doc = $word.Documents.Add()
    foreach ($source in (Get-RAToolsVbaSourceFiles -RepoRoot $repoRoot)) {
        [void]$doc.VBProject.VBComponents.Import($source.Path)
    }
    $components = $doc.VBProject.VBComponents
    # Only bypass interactive mode selection and completion dialogs in memory.
    $ui = $components.Item('Core_UI').CodeModule
    $start = $ui.ProcStartLine('ChooseBatchMode', 0)
    $count = $ui.ProcCountLines('ChooseBatchMode', 0)
    $ui.DeleteLines($start, $count)
    $ui.AddFromString("Public Function ChooseBatchMode(ByVal dialogTitle As String) As String`r`nChooseBatchMode = BATCH_MODE_CURRENT`r`nEnd Function")
    $clean = $components.Item('Macro_BatchAcceptAndClean').CodeModule
    $cleanText = $clean.Lines(1, $clean.CountOfLines)
    $cleanText = [regex]::Replace($cleanText, '(?m)^(\s*)MsgBox ', '$1Debug.Print ')
    $clean.DeleteLines(1, $clean.CountOfLines)
    $clean.AddFromString($cleanText)
    $manager = $components.Item('UI_ReviewToolbar').CodeModule
    $managerText = $manager.Lines(1, $manager.CountOfLines)
    $managerText = $managerText.Replace('If Len(failureText) > 0 Then MsgBox failureText, vbExclamation, REVIEW_TOOLBAR_CAPTION', "On Error GoTo 0`r`n    If Len(failureText) > 0 Then Err.Raise vbObjectError + 3001, , failureText")
    $manager.DeleteLines(1, $manager.CountOfLines)
    $manager.AddFromString($managerText)
    $probe = $components.Add(1)
    $probe.Name = 'ReviewActualActionsProbe'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private checks As Long
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 3010, , message
End Sub
Public Function RunActualReviewActions() As String
    Dim doc As Document
    Dim beforeText As String
    Dim beforeComments As Long
    On Error GoTo Failed
    Set doc = ThisDocument
    doc.Content.Text = "Selected text for real comment"
    doc.TrackRevisions = False
    doc.Activate
    doc.Range(0, 8).Select
    ShowReviewToolbar
    beforeText = doc.Content.Text
    beforeComments = doc.Comments.Count
    RunReviewToolbarAction "new_comment"
    Check doc.Comments.Count = beforeComments + 1, "Real native new comment"
    Check doc.Comments(doc.Comments.Count).Scope.Text = "Selected", "Real comment anchors original selection"
    Check Replace(doc.Content.Text, Chr$(5), "") = beforeText, "Comment preserves body text apart from Word comment reference marker"
    RunReviewToolbarAction "reviewer_filter"
    Check IsUserFormWindowVisible(REVIEWER_FILTER_CAPTION), "Real reviewer filter opens"
    Check frmReviewerFilter.FilteredDocumentKey <> "", "Reviewer filter binds a document"
    RunReviewToolbarAction "reviewer_filter"
    Check IsUserFormWindowVisible(REVIEWER_FILTER_CAPTION), "Existing reviewer filter remains interactive"
    Unload frmReviewerFilter
    doc.Activate
    RunReviewToolbarAction "track_changes"
    doc.Content.InsertAfter " tracked insertion"
    Check doc.Revisions.Count > 0, "Tracked edit fixture"
    Check doc.TrackRevisions, "Track enabled before clean"
    RunReviewToolbarAction "accept_revisions_comments"
    Check doc.Revisions.Count = 0, "Real clean accepts revisions"
    Check doc.Comments.Count = 0, "Real clean deletes comments"
    Check Not doc.TrackRevisions, "Real clean disables tracking"
    Check Not doc.Saved, "Current-document clean leaves pending edits unsaved"
    Check InStr(frmReviewToolbar.btnTrackChanges.Caption, "已关闭") > 0, "Toolbar refreshes after real clean"
    UnloadReviewToolbarOnQuit
    RunActualReviewActions = "PASS ReviewToolbar actual actions (" & checks & " assertions; mode selection, completion prompts and manager error presentation adapted in memory)"
    Exit Function
Failed:
    RunActualReviewActions = "FAIL actual actions after " & checks & ": " & Err.Description
End Function
'@)
    $result = [string]$word.Run('RunActualReviewActions')
    Write-Output $result
    if (-not $result.StartsWith('PASS')) { throw $result }
    Write-Output 'PASS ReviewToolbarActions.Tests'
}
finally {
    try {
        if ($null -ne $word) {
            try { $null = $word.Run('UnloadReviewToolbarOnQuit') } catch { }
            foreach ($document in @($word.Documents)) { $document.Close(0) }
        }
    }
    finally { Stop-RAToolsWordSession -Word $word }
}
