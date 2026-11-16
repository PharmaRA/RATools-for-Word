$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'scripts/RATools.Build.psm1') -Force
$testRoot = Join-Path $repoRoot ('backups/review-toolbar-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
$hostPath = Join-Path $testRoot 'review-host.dotm'
New-RAToolsDotmFromDirectory -SourceDirectory (Join-Path $repoRoot 'dotm') -DestinationPath $hostPath | Out-Null
function Replace-Procedure($Component, [string]$Name, [string]$Body) {
    $module = $Component.CodeModule
    $start = $module.ProcStartLine($Name, 0)
    $count = $module.ProcCountLines($Name, 0)
    $module.DeleteLines($start, $count)
    $module.AddFromString($Body)
}
$word = $null
try {
    $word = Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $hostDoc = $word.Documents.Open($hostPath, $false, $false, $false)
    foreach ($component in @($hostDoc.VBProject.VBComponents)) {
        if ([int]$component.Type -ne 100) { $hostDoc.VBProject.VBComponents.Remove($component) }
    }
    $hostDoc.Variables.Add('ReviewFixtureRoot', $testRoot) | Out-Null
    foreach ($source in (Get-RAToolsVbaSourceFiles -RepoRoot $repoRoot)) {
        [void]$hostDoc.VBProject.VBComponents.Import($source.Path)
    }
    $components = $hostDoc.VBProject.VBComponents
    $form = $components.Item('frmReviewToolbar')
    $form.CodeModule.AddFromString(@'
Public Sub ReviewClickProbe(ByVal buttonName As String)
    Select Case buttonName
        Case "btnTrackChanges": btnTrackChanges_Click
        Case "btnNewComment": btnNewComment_Click
        Case "btnReviewerFilter": btnReviewerFilter_Click
        Case "btnAcceptFormat": btnAcceptFormat_Click
        Case "btnDetectHighlights": btnDetectHighlights_Click
        Case "btnDetectClean": btnDetectClean_Click
        Case "btnClean": btnClean_Click
    End Select
End Sub
'@)
    $probe = $components.Add(1)
    $probe.Name = 'ReviewToolbarTestHost'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Route As String
Public BusyObserved As Boolean
Public ReentryBlocked As Boolean
Public ReopenPath As String
Public RaiseInRoute As Boolean
Private checks As Long

Public Sub RecordReviewRoute(ByVal name As String)
    Dim previousTrack As Boolean
    Route = name
    BusyObserved = Not frmReviewToolbar.btnTrackChanges.Enabled
    previousTrack = ActiveDocument.TrackRevisions
    RunReviewToolbarAction "track_changes"
    ReentryBlocked = (ActiveDocument.TrackRevisions = previousTrack)
    If RaiseInRoute Then Err.Raise 5, , "Expected route failure"
    If name = "accept_format_revisions" Then
        ActiveDocument.Close wdDoNotSaveChanges
        Documents.Open ReopenPath
    End If
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 3000, , message
End Sub

Private Function FormCount() As Long
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmReviewToolbar" Then FormCount = FormCount + 1
    Next item
End Function

Public Function RunReviewIntegration() As String
    Dim target As Document
    Dim other As Document
    Dim item As Object
    Dim buttonCount As Long
    Dim buttons As Variant
    Dim routes As Variant
    Dim index As Long
    Dim selectionStart As Long
    Dim selectionEnd As Long

    On Error GoTo Failed
    checks = 0
    RefreshReviewToolbarIfVisible
    Check FormCount() = 0, "Refresh must not load form"
    Set target = Documents.Add
    target.Content.Text = "Review selection fixture"
    ReopenPath = ThisDocument.Variables("ReviewFixtureRoot").Value & "\review-target.docx"
    target.SaveAs2 FileName:=ReopenPath, FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    target.TrackRevisions = False
    ShowReviewToolbar
    Check IsUserFormWindowVisible(REVIEW_TOOLBAR_CAPTION), "Show modeless form"
    Check GetUserFormWindowOwner(REVIEW_TOOLBAR_CAPTION) = 0, "Detach document owner"
    Check IsUserFormWindowOnTop(REVIEW_TOOLBAR_CAPTION), "Keep topmost"
    For Each item In frmReviewToolbar.Controls
        If TypeName(item) = "CommandButton" Then
            buttonCount = buttonCount + 1
            Check Not item.TakeFocusOnClick, "Keep text selection on click"
        End If
    Next item
    Check buttonCount = 7, "Exactly seven buttons"
    Check InStr(frmReviewToolbar.lblDocument.Caption, target.Name) > 0, "Current document label"
    frmReviewToolbar.ReviewClickProbe "btnTrackChanges"
    Check target.TrackRevisions, "Track toggle on current document"
    ShowQuickToolbar
    Check IsUserFormWindowVisible("RATools"), "Quick toolbar coexists"
    Check IsUserFormWindowVisible(REVIEW_TOOLBAR_CAPTION), "Review toolbar coexists"
    ShowReviewToolbar
    Check FormCount() = 1, "Single review instance"
    Set other = Documents.Add
    other.TrackRevisions = False
    RefreshReviewToolbarIfVisible
    Check InStr(frmReviewToolbar.lblDocument.Caption, other.Name) > 0, "Switch target document"
    frmReviewToolbar.ReviewClickProbe "btnTrackChanges"
    Check other.TrackRevisions, "Toggle new active target"
    other.Close wdDoNotSaveChanges
    target.Activate
    target.Range(2, 8).Select
    selectionStart = Selection.Start
    selectionEnd = Selection.End
    buttons = Array("btnNewComment", "btnReviewerFilter", "btnDetectHighlights", "btnDetectClean", "btnClean", "btnAcceptFormat")
    routes = Array("new_comment", "reviewer_filter", "detect_highlights", "detect_clean_versions", "accept_revisions_comments", "accept_format_revisions")
    For index = 0 To UBound(buttons)
        Route = ""
        frmReviewToolbar.ReviewClickProbe CStr(buttons(index))
        Check Route = routes(index), "Dispatch " & buttons(index)
        Check BusyObserved, "Busy " & buttons(index)
        Check ReentryBlocked, "Reentry " & buttons(index)
        Check frmReviewToolbar.btnTrackChanges.Enabled, "Restore controls " & buttons(index)
        If index = 0 Then Check Selection.Start = selectionStart And Selection.End = selectionEnd, "Selection preserved by comment route"
    Next index
    Check IsUserFormWindowVisible(REVIEW_TOOLBAR_CAPTION), "Survive document close and reopen"
    Check InStr(frmReviewToolbar.lblDocument.Caption, "review-target.docx") > 0, "Reopened target label"
    RaiseInRoute = True
    frmReviewToolbar.ReviewClickProbe "btnDetectHighlights"
    Check frmReviewToolbar.btnTrackChanges.Enabled, "Failure restores controls"
    RaiseInRoute = False
    frmReviewToolbar.ReviewClickProbe "btnDetectClean"
    Check Route = "detect_clean_versions", "Next action after failure"
    ToggleReviewToolbar Nothing
    Check Not IsUserFormWindowVisible(REVIEW_TOOLBAR_CAPTION), "Toggle hides"
    RefreshReviewToolbarIfVisible
    Check Not IsUserFormWindowVisible(REVIEW_TOOLBAR_CAPTION), "Refresh keeps hidden"
    ShowReviewToolbar
    frmReviewToolbar.Left = 123
    frmReviewToolbar.Top = 145
    UnloadReviewToolbarOnQuit
    Check FormCount() = 0, "Unload on quit"
    ShowReviewToolbar
    Check frmReviewToolbar.Left = 123 And frmReviewToolbar.Top = 145, "Independent position restoration"
    UnloadReviewToolbarOnQuit
    UnloadQuickToolbarOnQuit
    ActiveDocument.Close wdDoNotSaveChanges
    ThisDocument.Activate
    RunReviewIntegration = "PASS ReviewToolbar integration (" & checks & " assertions; modal routes stubbed in memory)"
    Exit Function
Failed:
    RunReviewIntegration = "FAIL after " & checks & " assertions: " & Err.Description
End Function

Public Function RunReviewNoDocuments() As String
    Dim doc As Document
    On Error GoTo Failed
    checks = 0
    Check Documents.Count = 0, "Zero-document fixture"
    ShowReviewToolbar
    Check Not frmReviewToolbar.btnTrackChanges.Enabled, "No-document actions disabled"
    RunReviewToolbarAction "track_changes"
    Check Documents.Count = 0, "No-document action creates nothing"
    Set doc = Documents.Add
    RefreshReviewToolbarIfVisible
    Check frmReviewToolbar.btnTrackChanges.Enabled, "New document enables actions"
    doc.Close wdDoNotSaveChanges
    RefreshReviewToolbarIfVisible
    Check Not frmReviewToolbar.btnTrackChanges.Enabled, "Last document close disables actions"
    UnloadReviewToolbarOnQuit
    Check FormCount() = 0, "No-document cleanup unloads"
    RunReviewNoDocuments = "PASS ReviewToolbar zero documents (" & checks & " assertions)"
    Exit Function
Failed:
    RunReviewNoDocuments = "FAIL no documents: " & Err.Description
End Function
'@)
    $routes = @(
        @('UI_QuickToolbarActions', 'ExecuteQuickToolbarMso', 'ByVal commandId As String', 'new_comment'),
        @('Macro_ReviewerFilter', 'Wrapper_ShowReviewerFilter', '', 'reviewer_filter'),
        @('Macro_AcceptFormatRevisions', 'Wrapper_AcceptFormatRevisions', '', 'accept_format_revisions'),
        @('Macro_BatchDetectHighlights', 'BatchDetectHighlights', '', 'detect_highlights'),
        @('Macro_BatchDetectCleanVersions', 'BatchDetectCleanVersions', '', 'detect_clean_versions'),
        @('Macro_BatchAcceptAndClean', 'BatchAcceptAndClean', '', 'accept_revisions_comments')
    )
    foreach ($route in $routes) {
        Replace-Procedure $components.Item($route[0]) $route[1] ("Public Sub " + $route[1] + '(' + $route[2] + ")`r`nRecordReviewRoute " + '"' + $route[3] + '"' + "`r`nEnd Sub")
    }
    $manager = $components.Item('UI_ReviewToolbar').CodeModule
    $managerText = $manager.Lines(1, $manager.CountOfLines)
    $managerText = $managerText.Replace('If Len(failureText) > 0 Then MsgBox failureText, vbExclamation, REVIEW_TOOLBAR_CAPTION', 'If Len(failureText) > 0 Then Debug.Print failureText')
    $manager.DeleteLines(1, $manager.CountOfLines)
    $manager.AddFromString($managerText)
    Write-Output 'STEP run integrated VBA scenario'
    $result = [string]$word.Run('RunReviewIntegration')
    Write-Output $result
    if (-not $result.StartsWith('PASS')) { throw $result }
    $hostDoc.Save()
    $hostDoc.Close(0)
    $hostDoc = $null
    [void]$word.AddIns.Add($hostPath, $true)
    Write-Output 'STEP run zero-document scenario'
    $result = [string]$word.Run('RunReviewNoDocuments')
    Write-Output $result
    if (-not $result.StartsWith('PASS')) { throw $result }
    Write-Output 'PASS ReviewToolbar.Tests'
}
finally {
    try {
        if ($null -ne $word) {
            try { $null = $word.Run('UnloadReviewToolbarOnQuit'); $null = $word.Run('UnloadQuickToolbarOnQuit') } catch { }
            foreach ($document in @($word.Documents)) { $document.Close(0) }
        }
    }
    finally {
        try { Stop-RAToolsWordSession -Word $word }
        finally { Move-RAToolsToRecycleBin -Path $testRoot -SafeRoot $repoRoot }
    }
}
