param([Parameter(Mandatory)][string]$PackagePath)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
$resolved=(Resolve-Path -LiteralPath $PackagePath).Path
function Normalize-Code([string]$text) {
    $lines=@($text -split '\r?\n' | Where-Object { $_ -notmatch '^Attribute ' -and $_.Trim().Length -gt 0 } | ForEach-Object {
        $line=$_.Trim()
        $cursor=0
        $normalized=[Text.StringBuilder]::new()
        foreach($literal in [regex]::Matches($line,'"(?:[^"]|"")*"')) {
            [void]$normalized.Append($line.Substring($cursor,$literal.Index-$cursor).ToLowerInvariant())
            [void]$normalized.Append($literal.Value)
            $cursor=$literal.Index+$literal.Length
        }
        [void]$normalized.Append($line.Substring($cursor).ToLowerInvariant())
        $normalized.ToString()
    })
    return $lines -join "`n"
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip=[IO.Compression.ZipFile]::OpenRead($resolved)
try {
    foreach($name in @('[Content_Types].xml','word/document.xml','word/vbaProject.bin','customUI/customUI14.xml')) {
        if($null -eq $zip.GetEntry($name)){throw "Missing required package part: $name"}
    }
} finally {$zip.Dispose()}
$word=$null;$doc=$null
try {
    $word=Start-RAToolsWordSession
    $word.AutomationSecurity=3
    $word.AddIns.Unload($false)
    $doc=$word.Documents.Open($resolved,$false,$true,$false)
    $sources=@(Get-RAToolsVbaSourceFiles -RepoRoot $repo)
    $checked=0
    foreach($source in $sources) {
        $name=[IO.Path]::GetFileNameWithoutExtension($source.Name)
        $component=$doc.VBProject.VBComponents.Item($name)
        if($null -eq $component){throw "Built package missing component $name"}
        if($source.Type -eq 'StandardModule') {
            $disk=[IO.File]::ReadAllText($source.Path,[Text.Encoding]::GetEncoding(936))
            $built=$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
            if((Normalize-Code $disk) -cne (Normalize-Code $built)){throw "Built module differs from source: $name"}
            $checked++
        }
    }
    if($doc.VBProject.VBComponents.Count -ne ($sources.Count+1)){throw 'Unexpected components in built package'}
    Write-Host "PASS built package opens, $checked standard modules match source, all $($sources.Count) source components present"
} finally {
    if($null -ne $doc){$doc.Close(0)}
    Stop-RAToolsWordSession -Word $word
}
Write-Host 'PASS BuiltPackage.Tests'
