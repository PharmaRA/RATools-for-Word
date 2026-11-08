$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
$root = Join-Path (Join-Path $repo 'backups') ('RATools_PdfSafety_' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($root)
Add-Type -AssemblyName Microsoft.VisualBasic
$word = $null; $hostDoc = $null
function Assert-True($condition,$message) { if (-not $condition) { throw $message } }
try {
    $word = Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $hostDoc = $word.Documents.Add()
    foreach ($name in @('Core_Files','Core_Session','Core_UI')) {
        [void]$hostDoc.VBProject.VBComponents.Import((Join-Path $repo "modules\$name.bas"))
    }
    $gbk = [Text.Encoding]::GetEncoding(936)
    $source = [IO.File]::ReadAllText((Join-Path $repo 'modules\Macro_BatchConvertWordToPDF.bas'),$gbk)
    $source = [regex]::Replace($source,'\bMsgBox\b','CapturePdfMessage')
    $modulePath = Join-Path $root 'Macro_BatchConvertWordToPDF.bas'
    [IO.File]::WriteAllText($modulePath,$source,$gbk)
    $component = $hostDoc.VBProject.VBComponents.Import($modulePath)
    foreach ($extension in @('doc','docx')) {
        $fixture = $word.Documents.Add()
        $fixture.Content.Text = "Unique source $extension"
        $format = 12
        if ($extension -eq 'doc') { $format = 0 }
        $fixture.SaveAs2([string](Join-Path $root "Report.$extension"),$format)
        $fixture.Close(0)
    }
    $existing = Join-Path $root 'Report.pdf'
    [IO.File]::WriteAllText($existing,'existing output must remain')
    $hash = (Get-FileHash -LiteralPath $existing).Hash
    $component.CodeModule.AddFromString(@"
Private Function CapturePdfMessage(ByVal message As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
End Function
Public Function PdfCollisionProbe() As String
    Dim session As ConvertSession, target As Document
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    session.updateType = 1
    Set target = Documents.Open("$root\Report.doc", Visible:=False)
    target.Activate
    ConvertActiveDocument session
    target.Close wdDoNotSaveChanges
    Set target = Documents.Open("$root\Report.docx", Visible:=False)
    target.Activate
    ConvertActiveDocument session
    target.Close wdDoNotSaveChanges
    PdfCollisionProbe = CStr(session.successCount) & "|" & session.processLog
End Function
"@)
    $hostDoc.Activate()
    $result = $word.Run('PdfCollisionProbe')
    Assert-True ($result -like '2|*') "Both sources must export: $result"
    Assert-True ((Get-FileHash -LiteralPath $existing).Hash -eq $hash) 'Existing PDF must never be overwritten'
    foreach ($name in @('Report_1.pdf','Report_2.pdf')) {
        $path = Join-Path $root $name
        Assert-True (Test-Path -LiteralPath $path) "Each source needs its own output: $name"
        $bytes = [IO.File]::ReadAllBytes($path)
        Assert-True ([Text.Encoding]::ASCII.GetString($bytes,0,5) -eq '%PDF-') 'Export must be a real PDF'
        Assert-True ($result.Contains($name)) 'Report must identify the actual output path'
    }
    Assert-True ((Get-FileHash (Join-Path $root 'Report_1.pdf')).Hash -ne (Get-FileHash (Join-Path $root 'Report_2.pdf')).Hash) 'Different source documents must keep distinct PDF contents'
    Write-Host 'PASS existing PDF retained and same-basename sources exported separately'
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'Report_3.pdf'))
    $code = $component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
    Assert-True ($code.Contains('fso.MoveFile tempPdf, pdfFileName')) 'Expected publication boundary for race injection'
    $code = $code.Replace('fso.MoveFile tempPdf, pdfFileName','RepairMoveWithRace fso, tempPdf, pdfFileName')
    $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
    $component.CodeModule.AddFromString($code)
    $component.CodeModule.AddFromString(@"
Private Sub RepairMoveWithRace(ByVal fso As Object, ByVal source As String, ByVal target As String)
    Static injected As Boolean
    Dim competitor As Object
    If Not injected Then
        injected = True
        Set competitor = fso.CreateTextFile(target, False)
        competitor.Write "concurrent writer"
        competitor.Close
    End If
    fso.MoveFile source, target
End Sub
Public Function PdfRaceProbe() As String
    Dim session As ConvertSession, target As Document
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    session.updateType = 1
    Set target = Documents.Open("$root\Report.docx", Visible:=False)
    target.Activate
    ConvertActiveDocument session
    target.Close wdDoNotSaveChanges
    PdfRaceProbe = CStr(session.successCount) & "|" & session.processLog
End Function
"@)
    $hostDoc.Activate()
    $race = $word.Run('PdfRaceProbe')
    Assert-True ($race -like '1|*' -and $race.Contains('Report_5.pdf')) 'Race and directory collision must choose the next available name'
    Assert-True ([IO.File]::ReadAllText((Join-Path $root 'Report_4.pdf')) -eq 'concurrent writer') 'Concurrent writer must never be overwritten'
    Assert-True (Test-Path -LiteralPath (Join-Path $root 'Report_3.pdf') -PathType Container) 'Existing directory must not be replaced'
    $raceBytes = [IO.File]::ReadAllBytes((Join-Path $root 'Report_5.pdf'))
    Assert-True ([Text.Encoding]::ASCII.GetString($raceBytes,0,5) -eq '%PDF-') 'Retried publication must retain valid PDF bytes'
    Write-Host 'PASS concurrent publication and directory collisions preserve other outputs'
    $component.CodeModule.AddFromString(@"
Public Function ExistingOpenPdfProbe() As String
    Dim session As ConvertSession
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    session.updateType = 1
    session.totalFileCount = 1
    ConvertOneFile session, "$root\Report.docx"
    ExistingOpenPdfProbe = CStr(session.successCount) & "|" & CStr(session.failCount) & "|" & session.processLog
End Function
"@)
    $existingDoc=$word.Documents.Open([string](Join-Path $root 'Report.docx'))
    $existingDoc.Content.InsertAfter(' unsaved user edit')
    $lockedField=$existingDoc.Fields.Add($existingDoc.Range(0,0),33,'PAGE')
    $lockedField.Locked=$true
    $unlockedField=$existingDoc.Fields.Add($existingDoc.Range(0,0),33,'PAGE')
    $unlockedField.Locked=$false
    $pendingText=$existingDoc.Content.Text
    $countBefore=$word.Documents.Count
    $hostDoc.Activate()
    $openResult=$word.Run('ExistingOpenPdfProbe')
    Assert-True ($openResult -like '1|0|*') 'Already-open document must still export successfully'
    Assert-True ($word.Documents.Count -eq $countBefore) 'Batch export must not close a user-owned open document'
    Assert-True ($existingDoc.Content.Text -eq $pendingText -and -not $existingDoc.Saved) 'Batch export must preserve unsaved user edits'
    Assert-True ($lockedField.Locked -and -not $unlockedField.Locked) 'Batch export must restore original user field locks'
    $code=$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
    $code=[regex]::Replace($code,'(?ms)^Private Sub SafeExportAsPDF\(.*?^End Sub',@'
Private Sub SafeExportAsPDF(doc As Document, ByRef pdfFileName As String)
    Err.Raise vbObjectError + 2502, , "Injected existing document export failure"
End Sub
'@)
    $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
    $component.CodeModule.AddFromString($code)
    $hostDoc.Activate()
    $failedOpenResult=$word.Run('ExistingOpenPdfProbe')
    Assert-True ($failedOpenResult -like '0|1|*Injected existing document export failure*') 'Failed user document export must report original error'
    Assert-True ($word.Documents.Count -eq $countBefore -and $existingDoc.Content.Text -eq $pendingText) 'Failed batch export must retain the user document and text'
    Assert-True ($lockedField.Locked -and -not $unlockedField.Locked -and -not $existingDoc.Saved) 'Failed batch export must restore locks and retain pending edits'
    $existingDoc.Close(0)
    Write-Host 'PASS batch export preserves already-open unsaved document'
} finally {
    if ($null -ne $word) {
        try { for ($i=$word.Documents.Count;$i -ge 1;$i--) { $word.Documents.Item($i).Close(0) } } catch {}
    }
    Stop-RAToolsWordSession -Word $word
    try { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($root,'OnlyErrorDialogs','SendToRecycleBin') }
    catch { Write-Warning "Test directory retained: $root" }
}
Write-Host 'PASS PdfOutputSafety.Tests'
