$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
Add-Type -AssemblyName Microsoft.VisualBasic
$root = Join-Path (Join-Path $repo 'backups') ('RATools_SaveRepair_' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($root)
$word = $null; $hostDoc = $null
function Assert-True($condition,$message) { if (-not $condition) { throw $message } }
try {
    $word = Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $hostDoc = $word.Documents.Add()
    foreach ($name in @('Core_Files','Core_UI','Core_Session','Macro_BatchAcceptAndClean','Macro_BatchSetMargins')) {
        [void]$hostDoc.VBProject.VBComponents.Import((Join-Path $repo "modules\$name.bas"))
    }
    $events = $hostDoc.VBProject.VBComponents.Add(2)
    $events.Name = 'CancelBatchSave'
    $events.CodeModule.AddFromString(@'
Public WithEvents App As Word.Application
Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    Cancel = True
End Sub
'@)
    foreach ($module in @('Macro_BatchAcceptAndClean','Macro_BatchSetMargins')) {
        Write-Host "Creating fixture for $module"
        $path = Join-Path $root ($module + '.docx')
        $fixture = $word.Documents.Add()
        $fixture.Content.Text = 'Original document'
        $fixture.PageSetup.TopMargin = 100
        $fixture.TrackRevisions = $true
        $fixture.Content.InsertAfter(' tracked addition')
        $fixture.SaveAs2([string]$path,12)
        Write-Host "Saved fixture for $module"
        $fixture.Close(0)
        $before = (Get-FileHash -LiteralPath $path).Hash
        $component = $hostDoc.VBProject.VBComponents.Item($module)
        $probeName = 'SaveFailureProbe_' + $module
        $component.CodeModule.AddFromString(@"
Public Function $probeName() As Boolean
    Dim events As CancelBatchSave
    Set events = New CancelBatchSave
    Set events.App = Application
    $probeName = ProcessFile("$path")
    Set events.App = Nothing
End Function
"@)
        $hostDoc.Activate()
        Write-Host "Running cancelled save for $module"
        Assert-True ($word.Run($probeName) -eq $false) "$module must count a cancelled save as failure"
        $retained = $word.Documents.Item([string]$path)
        Assert-True (-not $retained.Saved) "$module must retain pending changes"
        if ($word.Visible) { Assert-True ($retained.ActiveWindow.Visible) 'Interactive Word must reveal failed background document' }
        $retained.Close(0)
        Assert-True ((Get-FileHash -LiteralPath $path).Hash -eq $before) 'Cancelled save must keep disk original'
        Write-Host "PASS $module reports failure and retains recoverable document"
        if ($module -eq 'Macro_BatchAcceptAndClean') {
            $existing = $word.Documents.Open([string]$path)
            $existing.Content.InsertAfter(' user pending edit')
            $pendingText = $existing.Content.Text
            $hostDoc.Activate()
            Assert-True ($word.Run($probeName) -eq $false) 'Already-open document must report cancelled save as failure'
            Assert-True ($word.Documents.Count -eq 2 -and -not $existing.Saved) 'Already-open document must stay open and dirty'
            Assert-True ($existing.Content.Text -eq $pendingText) 'Already-open pending text must survive'
            $existing.Close(0)
            Assert-True ((Get-FileHash -LiteralPath $path).Hash -eq $before) 'Already-open save cancellation must keep disk original'
            Write-Host 'PASS already-open document retains pending edits and reports failure'
        }
        $successName = 'SuccessfulSaveProbe_' + $module
        $component.CodeModule.AddFromString(@"
Public Function $successName() As Boolean
    $successName = ProcessFile("$path")
End Function
"@)
        $hostDoc.Activate()
        Assert-True ($word.Run($successName) -eq $true) "$module normal save must succeed"
        Assert-True ($word.Documents.Count -eq 1) 'Owned document must close after successful save'
        $check = $word.Documents.Open([string]$path)
        if ($module -eq 'Macro_BatchAcceptAndClean') { Assert-True ($check.Revisions.Count -eq 0) 'Cleaned document must be saved' }
        else { Assert-True ([Math]::Abs($check.PageSetup.TopMargin - 72) -lt 0.1) 'Margins must be saved' }
        $check.Close(0)
        Write-Host "PASS $module normal save persists changes"
    }
    $hostDoc.VBProject.VBComponents.Item('Core_Session').CodeModule.AddFromString(@'
Public Function CancelledCloseProbe() As Boolean
    Dim target As Document, events As CancelBatchSave, countBefore As Long
    Set target = Documents.Add
    target.Content.Text = "Unsaved edits must survive"
    Set events = New CancelBatchSave
    Set events.App = Application
    countBefore = Documents.Count
    CloseDocumentQuietly target, True
    Set events.App = Nothing
    CancelledCloseProbe = (Documents.Count = countBefore)
End Function
'@)
    Assert-True ($word.Run('CancelledCloseProbe') -eq $true) 'Cancelled save must not close or discard the unsaved document'
    $pending = $word.ActiveDocument
    Assert-True ($pending.Content.Text -like 'Unsaved edits must survive*') 'Pending contents must survive'
    $pending.Close(0)
    Write-Host 'PASS cancelled save keeps pending contents'
    $closeEvents = $hostDoc.VBProject.VBComponents.Add(2)
    $closeEvents.Name = 'CancelBatchClose'
    $closeEvents.CodeModule.AddFromString(@'
Public WithEvents App As Word.Application
Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    Cancel = True
End Sub
'@)
    $hostDoc.VBProject.VBComponents.Item('Core_Session').CodeModule.AddFromString(@'
Public Function CancelledHiddenCloseProbe() As Boolean
    Dim target As Document, events As CancelBatchClose, countBefore As Long
    Dim closed As Boolean
    Set target = Documents.Add(Visible:=False)
    Set events = New CancelBatchClose
    Set events.App = Application
    countBefore = Documents.Count
    closed = CloseDocumentQuietly(target, False)
    Set events.App = Nothing
    CancelledHiddenCloseProbe = Not closed And Documents.Count = countBefore
    If Application.Visible Then CancelledHiddenCloseProbe = CancelledHiddenCloseProbe And target.ActiveWindow.Visible
    target.Close wdDoNotSaveChanges
End Function
'@)
    $hostDoc.Activate()
    Assert-True ($word.Run('CancelledHiddenCloseProbe') -eq $true) 'Cancelled close must report failure and reveal the retained document'
    Write-Host 'PASS cancelled close retains document without forcing hidden automation visible'
    $marginModule=$hostDoc.VBProject.VBComponents.Item('Macro_BatchSetMargins')
    $marginCode=$marginModule.CodeModule.Lines(1,$marginModule.CodeModule.CountOfLines)
    $marginFault=[regex]::Replace($marginCode,'(?ms)^Private Function SetMarginsSafely\(.*?^End Function',@'
Private Function SetMarginsSafely(doc As Document) As Boolean
    doc.PageSetup.LeftMargin = 80
    SetMarginsSafely = False
End Function
'@)
    Assert-True ($marginFault -ne $marginCode) 'Margin failure must replace the actual setter boundary'
    $marginModule.CodeModule.DeleteLines(1,$marginModule.CodeModule.CountOfLines)
    $marginModule.CodeModule.AddFromString($marginFault)
    $marginModule.CodeModule.AddFromString(@"
Public Function MarginPartialFailureProbe() As String
    On Error GoTo Failed
    MarginPartialFailureProbe = CStr(ProcessFile("$path"))
    Exit Function
Failed:
    MarginPartialFailureProbe = "ERROR:" & Err.Description
End Function
"@)
    $hostDoc.Activate()
    Assert-True ($word.Run('MarginPartialFailureProbe') -eq 'False') 'Partial margin failure must return false without throwing'
    Assert-True (-not $word.Visible) 'Partial margin failure must not expose hidden automation'
    $partial=$word.Documents.Item([string]$path)
    Assert-True (-not $partial.Saved) 'Partial margin changes must stay recoverable'
    $partial.Close(0)
    Write-Host 'PASS partial margin failure retains changes without changing automation visibility'
    $sessionModule=$hostDoc.VBProject.VBComponents.Item('Core_Session')
    $sessionCode=$sessionModule.CodeModule.Lines(1,$sessionModule.CodeModule.CountOfLines)
    $sessionModule.CodeModule.DeleteLines(1,$sessionModule.CodeModule.CountOfLines)
    $faultCode=[regex]::Replace($sessionCode,'(?im)^\s*doc\.Save\s*$', '    Err.Raise vbObjectError + 2501, , "Injected persistence failure"')
    Assert-True ($faultCode -ne $sessionCode) 'Persistence fault must replace the actual VBA save call'
    $sessionModule.CodeModule.AddFromString($faultCode)
    $sessionModule.CodeModule.AddFromString(@'
Public Function PersistenceErrorProbe() As Boolean
    Dim target As Document, errorText As String, saved As Boolean, before As Long
    Set target = Documents.Add(Visible:=False)
    target.Content.Text = "Retain on I/O exception"
    before = Documents.Count
    saved = CloseDocumentQuietly(target, True, errorText)
    PersistenceErrorProbe = Not saved And Documents.Count = before And Not target.Saved _
        And errorText = "Injected persistence failure" And target.Content.Text = "Retain on I/O exception" & vbCr
    target.Close wdDoNotSaveChanges
End Function
'@)
    $hostDoc.Activate()
    Assert-True ($word.Run('PersistenceErrorProbe')) 'Save exception must return failure, preserve error and pending content'
    Assert-True (-not $word.Visible) 'Failure recovery must not force a hidden automation instance visible'
    Write-Host 'PASS persistence exception keeps pending content and propagates failure'
} catch {
    Write-Host ('TEST FAILURE: ' + $_.Exception.Message + ' at ' + $_.InvocationInfo.PositionMessage)
    throw
} finally {
    if ($null -ne $word) {
        try {
            for ($documentIndex = $word.Documents.Count; $documentIndex -ge 1; $documentIndex--) {
                $word.Documents.Item($documentIndex).Close(0)
            }
        } catch { Write-Warning ('Document cleanup failed: ' + $_.Exception.Message) }
    }
    Stop-RAToolsWordSession -Word $word
    try { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($root,'OnlyErrorDialogs','SendToRecycleBin') }
    catch { Write-Warning "Test directory retained: $root" }
}
Write-Host 'PASS BatchSaveSafety.Tests'
