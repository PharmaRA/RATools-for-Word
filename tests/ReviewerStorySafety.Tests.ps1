$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
$word=$null;$doc=$null
try {
    $word=Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    foreach($case in @('Header','Footer','Empty')) {
        $doc=$word.Documents.Add()
        [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'class_modules\clsReviewerItem.cls'))
        [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'userforms\frmReviewerFilter.frm'))
        $module=$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Macro_ReviewerFilter.bas'))
        $module.CodeModule.AddFromString(@'
Public Function StoryReviewerProbe() As Long
    Dim names() As String, counts() As Long, total As Long
    names = CollectReviewerNames(ActiveDocument, counts, total)
    StoryReviewerProbe = total
End Function
'@)
        $doc.Content.Text='Unchanged body'
        $doc.TrackRevisions=$true
        if($case -eq 'Header'){$doc.Sections.Item(1).Headers.Item(1).Range.Text='Only header change'}
        if($case -eq 'Footer'){$doc.Sections.Item(1).Footers.Item(1).Range.Text='Only footer change'}
        if($doc.Revisions.Count -ne 0 -or $doc.Comments.Count -ne 0){throw 'Fixture must have no main-story revisions or comments'}
        $expected=1
        if($case -eq 'Empty'){$expected=0}
        $actual=$word.Run('StoryReviewerProbe')
        if($actual -ne $expected){throw "$case must return $expected reviewer(s), actual $actual"}
        Write-Host "PASS $case reviewer collection"
        $doc.Close(0);$doc=$null
    }
} finally {
    if($null -ne $doc){$doc.Close(0)}
    Stop-RAToolsWordSession -Word $word
}
Write-Host 'PASS ReviewerStorySafety.Tests'
