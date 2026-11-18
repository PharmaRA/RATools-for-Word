param([Parameter(Mandatory)][string]$PackagePath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'BuiltPackage.Tests.ps1') -PackagePath $PackagePath
$repoRoot = Split-Path -Parent $PSScriptRoot
$resolvedPackage = (Resolve-Path -LiteralPath $PackagePath).Path
$layout = Import-PowerShellDataFile (Join-Path $repoRoot 'scripts/WorkflowToolbarLayout.psd1')
$packageWord = $null
$packageDocument = $null
$checkedComponents = 0
$checkedControls = 0
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('RATools_WorkflowPackage_' + [guid]::NewGuid().ToString('N'))
try {
    [void][IO.Directory]::CreateDirectory($fixtureRoot)
    $fixturePath = Join-Path $fixtureRoot 'workflow-package.dotm'
    Copy-Item -LiteralPath $resolvedPackage -Destination $fixturePath
    $packageWord = Start-RAToolsWordSession
    $packageWord.AutomationSecurity = 3
    $packageWord.AddIns.Unload($false)
    $packageDocument = $packageWord.Documents.Open($fixturePath, $false, $false, $false)
    $gbk = [Text.Encoding]::GetEncoding(936)
    foreach ($relative in @('modules/UI_WorkflowToolbars.bas','class_modules/clsAppEvents.cls','userforms/frmWritingToolbar.frm','userforms/frmFinalToolbar.frm','userforms/frmReferencesToolbar.frm')) {
        $source = [IO.File]::ReadAllText((Join-Path $repoRoot $relative), $gbk)
        $source = $source.Substring($source.IndexOf('Option Explicit'))
        $name = [IO.Path]::GetFileNameWithoutExtension($relative)
        $component = $packageDocument.VBProject.VBComponents.Item($name)
        $built = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
        if ((Normalize-Code $source) -cne (Normalize-Code $built)) { throw ('Built workflow component differs: ' + $name) }
        $checkedComponents++
    }
    foreach ($kind in @('Writing','Final','References')) {
        $component = $packageDocument.VBProject.VBComponents.Item('frm' + $kind + 'Toolbar')
        $exportPath = Join-Path $fixtureRoot ('frm' + $kind + 'Toolbar.frm')
        $component.Export($exportPath)
        $exported = [IO.File]::ReadAllText($exportPath, $gbk)
        if ($exported -notmatch 'ShowModal\s*=\s*0') { throw ('Built window is modal: ' + $kind) }
        $designer = $component.Designer
        if ($null -eq $designer) { throw ('Built form designer unavailable after export: ' + $kind) }
        $expectedCount = @{Writing=31;Final=21;References=20}[$kind]
        if ($designer.Controls.Count -ne $expectedCount) { throw ('Built control count differs: ' + $kind) }
        if ($designer.Controls.Item('cboToolbars').Tag -ne '工具窗' -or $designer.Controls.Item('cboToolbars').Style -ne 2) { throw ('Built window menu differs: ' + $kind) }
        $index = 0
        foreach ($row in $layout[$kind].Rows) {
            for ($column = 1; $column -lt $row.Count; $column++) {
                $parts = $row[$column] -split '\|', 2
                $isMenu = $parts[1].StartsWith('menu:')
                $name = if ($isMenu) { 'cbo' + $parts[1].Substring(5) } else { 'btn' + $index }
                $control = $designer.Controls.Item($name)
                $expectedTag = if ($isMenu) { $parts[0] } else { $parts[1] }
                if ($control.Tag -cne $expectedTag -or -not $control.ControlTipText) { throw ('Built action metadata differs: ' + $kind + '/' + $name) }
                if (-not $isMenu -and ($control.Caption -cne $parts[0] -or $control.TakeFocusOnClick)) { throw ('Built button differs: ' + $kind + '/' + $name) }
                $checkedControls++
                $index++
            }
        }
    }
    $archive = [IO.Compression.ZipFile]::OpenRead($resolvedPackage)
    try {
        $reader = [IO.StreamReader]::new($archive.GetEntry('customUI/customUI14.xml').Open())
        try { $ribbon = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        $sourceRibbon = [xml]([IO.File]::ReadAllText((Join-Path $repoRoot 'dotm/customUI/customUI14.xml'), [Text.Encoding]::UTF8))
        $builtMenu = $ribbon.SelectSingleNode('//*[@id="menuWorkflowToolbars"]')
        $sourceMenu = $sourceRibbon.SelectSingleNode('//*[@id="menuWorkflowToolbars"]')
        if ($null -eq $builtMenu -or $builtMenu.OuterXml -cne $sourceMenu.OuterXml) { throw 'Built unified Ribbon menu differs from source.' }
        foreach ($legacy in @('btnQuickToolbar','btnReviewToolbar')) {
            if ($null -eq $ribbon.SelectSingleNode('//*[@id="' + $legacy + '"]')) { throw ('Built legacy window entry missing: ' + $legacy) }
        }
    } finally { $archive.Dispose() }
    Write-Output ('PASS WorkflowToolbarPackage (' + $checkedComponents + ' direct components match, ' + $checkedControls + ' action controls, three modeless windows and unified Ribbon)')
}
finally {
    try { if ($null -ne $packageDocument) { $packageDocument.Close(0) } }
    finally {
        try { Stop-RAToolsWordSession -Word $packageWord }
        finally { Move-RAToolsToRecycleBin -Path $fixtureRoot -SafeRoot ([IO.Path]::GetTempPath()) }
    }
}
