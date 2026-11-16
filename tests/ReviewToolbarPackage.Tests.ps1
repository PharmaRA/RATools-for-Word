param([Parameter(Mandatory)][string]$PackagePath)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'scripts/RATools.Build.psm1') -Force
$resolved = (Resolve-Path -LiteralPath $PackagePath).Path
$fixtureRoot = Join-Path $repoRoot ('backups/review-package-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixtureRoot)
$word = $null
$document = $null
try {
$fixturePath = Join-Path $fixtureRoot 'review-package.dotm'
Copy-Item -LiteralPath $resolved -Destination $fixturePath
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::OpenRead($resolved)
try {
    $entry = $archive.GetEntry('customUI/customUI14.xml')
    $reader = [IO.StreamReader]::new($entry.Open())
    try { [xml]$ribbon = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $button = $ribbon.SelectSingleNode('//*[local-name()="button" and @id="btnReviewToolbar"]')
    if ($null -eq $button -or $button.onAction -ne 'ToggleReviewToolbar') { throw 'Built review Ribbon callback is missing.' }
    if ($null -ne $ribbon.SelectSingleNode('//*[@id="btnReviewerFilter" or @id="btnAcceptFormatRevisions"]')) {
        throw 'Obsolete standalone review buttons remain in the package.'
    }
    foreach ($native in @('ReviewNewComment', 'ReviewTrackChanges')) {
        if ($null -eq $ribbon.SelectSingleNode('//*[@idMso="' + $native + '"]')) { throw "Native Ribbon entry missing: $native" }
    }
}
finally { $archive.Dispose() }
function Normalize-Code([string]$Code) {
    $codeLines = $Code -split '\r?\n' | Where-Object { $_ -notmatch '^Attribute ' -and $_.Trim() }
    return ($codeLines | ForEach-Object { $_.Trim() }) -join "`n"
}
    $word = Start-RAToolsWordSession
    $word.AutomationSecurity = 3
    $word.AddIns.Unload($false)
    $document = $word.Documents.Open($fixturePath, $false, $false, $false)
    $gbk = [Text.Encoding]::GetEncoding(936)
    foreach ($relative in @('userforms/frmReviewToolbar.frm', 'class_modules/clsAppEvents.cls')) {
        $source = [IO.File]::ReadAllText((Join-Path $repoRoot $relative), $gbk)
        $source = $source.Substring($source.IndexOf('Option Explicit'))
        $name = [IO.Path]::GetFileNameWithoutExtension($relative)
        $component = $document.VBProject.VBComponents.Item($name)
        $built = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
        if ((Normalize-Code $source) -ine (Normalize-Code $built)) { throw "Built component differs: $name" }
    }
    $form = $document.VBProject.VBComponents.Item('frmReviewToolbar')
    Write-Output ('Built review component: ' + $form.Name + '; type=' + $form.Type)
    $exportPath = Join-Path $fixtureRoot 'frmReviewToolbar.frm'
    $form.Export($exportPath)
    $exported = [IO.File]::ReadAllText($exportPath, $gbk)
    if ($exported -notmatch 'ShowModal\s*=\s*0') { throw 'Built form must be modeless.' }
    $designer = $form.Designer
    $buttons = @($designer.Controls | Where-Object { $_.Name -like 'btn*' })
    if ($buttons.Count -ne 7) { throw 'Built form must have seven action controls.' }
    foreach ($name in @('btnTrackChanges','btnNewComment','btnReviewerFilter','btnAcceptFormat','btnDetectHighlights','btnDetectClean','btnClean')) {
        $control = $designer.Controls.Item($name)
        if (-not $control.Caption -or -not $control.ControlTipText) { throw "Missing visible text or tip: $name" }
        if ($control.TakeFocusOnClick) { throw "Button steals selection focus: $name" }
    }
    Write-Output 'PASS ReviewToolbarPackage.Tests (Ribbon migration, modeless form, seven controls, form and event source match)'
}
finally {
    try {
        if ($null -ne $document) { $document.Close(0) }
    }
    finally {
        try { Stop-RAToolsWordSession -Word $word }
        finally { Move-RAToolsToRecycleBin -Path $fixtureRoot -SafeRoot $repoRoot }
    }
}
