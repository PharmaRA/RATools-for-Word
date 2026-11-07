$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
Add-Type -AssemblyName Microsoft.VisualBasic
$root = Join-Path (Join-Path $repo 'backups') ('RATools_FieldState_' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($root)
$word = $null; $doc = $null
function Assert-True($condition,$message) { if (-not $condition) { throw $message } }
try {
    $word = Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $doc = $word.Documents.Add()
    foreach ($name in @('Core_Files','Core_Session','Core_UI')) {
        [void]$doc.VBProject.VBComponents.Import((Join-Path $repo "modules\$name.bas"))
    }
    $encoding = [Text.Encoding]::GetEncoding(936)
    $source = [IO.File]::ReadAllText((Join-Path $repo 'modules\Macro_BatchConvertWordToPDF.bas'),$encoding)
    $source = $source.Replace('Option Explicit',"Option Explicit`r`nPrivate repairMessage As String")
    $source = [regex]::Replace($source,'\bMsgBox\b','CaptureFieldMessage')
    $modulePath = Join-Path $root 'Macro_BatchConvertWordToPDF.bas'
    [IO.File]::WriteAllText($modulePath,$source,$encoding)
    $component = $doc.VBProject.VBComponents.Import($modulePath)
    $component.CodeModule.AddFromString(@'
Private Function CaptureFieldMessage(ByVal message As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    repairMessage = message
End Function
Public Function FieldStateExportProbe() As Long
    Dim session As ConvertSession
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    session.updateType = 1
    repairMessage = ""
    ConvertActiveDocument session
    FieldStateExportProbe = session.successCount * 100 + session.failCount
End Function
Public Function FieldStateErrorProbe() As String
    FieldStateErrorProbe = repairMessage
End Function
'@)
    $doc.Content.Text = 'Body fields: '
    $first = $doc.Fields.Add($doc.Range(0,0),33,'PAGE')
    $position = $doc.Content.End-1
    $second = $doc.Fields.Add($doc.Range($position,$position),33,'PAGE')
    $header = $doc.Sections.Item(1).Headers.Item(1).Range
    $headerField = $header.Fields.Add($header,33,'PAGE')
    $shape = $doc.Shapes.AddTextbox(1,50,50,150,80)
    $shape.TextFrame.TextRange.Text = 'Shape '
    $shapeField = $shape.TextFrame.TextRange.Fields.Add($shape.TextFrame.TextRange,33,'PAGE')
    $expected = @(
        @{ Field=$first; Locked=$true; Name='locked body' },
        @{ Field=$second; Locked=$false; Name='unlocked body' },
        @{ Field=$headerField; Locked=$true; Name='locked header' },
        @{ Field=$shapeField; Locked=$false; Name='unlocked shape' }
    )
    foreach ($item in $expected) { $item.Field.Locked = $item.Locked }
    $doc.SaveAs2([string](Join-Path $root 'FieldState.docm'),13)
    Assert-True ($word.Run('FieldStateExportProbe') -eq 100) 'Real PDF export must succeed'
    foreach ($item in $expected) { Assert-True ($item.Field.Locked -eq $item.Locked) "Successful export must restore $($item.Name)" }
    Write-Host 'PASS successful real PDF export restores mixed field states'
    $code = $component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
    $code = [regex]::Replace($code,'(?ms)^Private Sub SafeExportAsPDF\(.*?^End Sub', @'
Private Sub SafeExportAsPDF(doc As Document, ByRef pdfFileName As String)
    Err.Raise vbObjectError + 2401, , "Injected export failure"
End Sub
'@)
    $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
    $component.CodeModule.AddFromString($code)
    Assert-True ($word.Run('FieldStateExportProbe') -eq 1) 'Failed export must count as failure'
    foreach ($item in $expected) { Assert-True ($item.Field.Locked -eq $item.Locked) "Failed export must restore $($item.Name)" }
    Assert-True ($word.Run('FieldStateErrorProbe') -like '*Injected export failure*') 'Restoration must not hide the original export error'
    Write-Host 'PASS failed export restores mixed field states and preserves error'
    $code = $component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
    $code = [regex]::Replace($code,'(?ms)^Private Sub SafeExportAsPDF\(.*?^End Sub', @'
Private Sub SafeExportAsPDF(doc As Document, ByRef pdfFileName As String)
    Err.Raise 5124, , "Injected mapped export error"
End Sub
'@)
    $code = [regex]::Replace($code,'(?ms)(^Private Function RestoreFieldLockStates\(.*?)(^End Function)', '$1    RestoreFieldLockStates = False' + "`r`n" + '$2')
    $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
    $component.CodeModule.AddFromString($code)
    Assert-True ($word.Run('FieldStateExportProbe') -eq 1) 'Combined export and restoration failures must count as failure'
    $restoreWarning = -join [char[]]@(0x9501,0x5b9a,0x72b6,0x6001,0x65e0,0x6cd5,0x6062,0x590d)
    Assert-True ($word.Run('FieldStateErrorProbe').Contains($restoreWarning)) 'Friendly error mapping must preserve the restoration warning'
    Write-Host 'PASS mapped export error retains partial restoration warning'
} finally {
    if ($null -ne $doc) { try { $doc.Close(0) } catch {} }
    Stop-RAToolsWordSession -Word $word
    try { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($root,'OnlyErrorDialogs','SendToRecycleBin') }
    catch { Write-Warning "Test directory retained: $root" }
}
Write-Host 'PASS PdfFieldState.Tests'
