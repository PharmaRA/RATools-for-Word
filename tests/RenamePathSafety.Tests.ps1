$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
Add-Type -AssemblyName Microsoft.VisualBasic
$root=Join-Path (Join-Path $repo 'backups') ('RATools_RenameSafety_'+[guid]::NewGuid().ToString('N'))
$data=Join-Path $root 'files'
[void][IO.Directory]::CreateDirectory($data)
[IO.File]::WriteAllText((Join-Path $data 'Report.DOCX'),'case-only original')
[IO.File]::WriteAllText((Join-Path $data 'A B.DOCX'),'new cleaned source')
[IO.File]::WriteAllText((Join-Path $data 'a-b.docx'),'existing conflicting target')
$word=$null;$doc=$null
try {
    $encoding=[Text.Encoding]::GetEncoding(936)
    $source=[IO.File]::ReadAllText((Join-Path $repo 'modules\Macro_BatchRenameFiles.bas'),$encoding)
    $source=[regex]::Replace($source,'\bMsgBox\b','ChooseRenameTestMode')
    $source=[regex]::Replace($source,'PickFolder\("[^"]*"\)',('"'+$data+'"'))
    $source+=@'

Private Function ChooseRenameTestMode(ByVal message As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    ChooseRenameTestMode = vbYes
End Function
'@
    $modulePath=Join-Path $root 'Macro_BatchRenameFiles.bas'
    [IO.File]::WriteAllText($modulePath,$source,$encoding)
    $word=Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $doc=$word.Documents.Add()
    [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Core_Session.bas'))
    [void]$doc.VBProject.VBComponents.Import($modulePath)
    $word.Run('BatchRenameFiles')
    $names=@(Get-ChildItem -LiteralPath $data -File | ForEach-Object {$_.Name})
    if($names -cnotcontains 'report.docx'){throw "Case-only rename must produce report.docx, actual $($names -join ', ')"}
    if($names.Count -ne 3){throw 'Rename must preserve file count'}
    if([IO.File]::ReadAllText((Join-Path $data 'report.docx')) -ne 'case-only original'){throw 'Case rename must preserve contents'}
    if([IO.File]::ReadAllText((Join-Path $data 'a-b.docx')) -ne 'existing conflicting target'){throw 'True conflict must not overwrite existing target'}
    if([IO.File]::ReadAllText((Join-Path $data 'a-b_1.docx')) -ne 'new cleaned source'){throw 'True conflict must use a suffix'}
    Write-Host 'PASS case-only rename and true collision preserve names and contents'
} finally {
    if($null -ne $doc){$doc.Close(0)}
    Stop-RAToolsWordSession -Word $word
    Move-RAToolsToRecycleBin $root -SafeRoot (Join-Path $repo 'backups')
}
Write-Host 'PASS RenamePathSafety.Tests'
