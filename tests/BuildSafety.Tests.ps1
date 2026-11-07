$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = Join-Path (Join-Path $repo 'backups') ('RATools_BuildSafety_' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($root)
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName Microsoft.VisualBasic
function Assert-True($condition,$message) { if (-not $condition) { throw $message } }
function Expect-Failure([scriptblock]$action) {
    try { & $action | Out-Null } catch { return }
    throw 'Expected operation to fail'
}
function New-FixturePackage($path, [bool]$complete = $true) {
    $zip = [IO.Compression.ZipFile]::Open($path, 'Create')
    try {
        $entries = @('[Content_Types].xml','word/document.xml')
        if ($complete) { $entries += 'word/vbaProject.bin' }
        foreach ($name in $entries) {
            $writer = [IO.StreamWriter]::new($zip.CreateEntry($name).Open())
            try { $writer.Write('new package data') } finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose() }
}
try {
    $source = [IO.File]::ReadAllText((Join-Path $repo 'scripts\RATools.Build.psm1'))
    # Baseline reproduction substitutes only physical cleanup, never validation,
    # extraction or replacement decisions. No test may physically delete files.
    $source = $source.Replace('Remove-Item -LiteralPath', 'Move-TestToRecycleBin -LiteralPath')
    $safety = @'
function Move-TestToRecycleBin {
    param([string]$LiteralPath,[switch]$Recurse,[switch]$Force)
    $path = [IO.Path]::GetFullPath($LiteralPath)
    if (-not $path.StartsWith($script:repairFixtureRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe baseline cleanup' }
    if ([IO.Directory]::Exists($path)) {
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($path,'OnlyErrorDialogs','SendToRecycleBin')
    } else {
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($path,'OnlyErrorDialogs','SendToRecycleBin')
    }
}
'@
    $module = New-Module -Name RepairBuildUnderTest -ScriptBlock ([scriptblock]::Create($safety + "`n" + $source))
    Import-Module $module -Force
    & $module { param($path) $script:repairFixtureRoot = $path } $root
    $target = Join-Path $root 'dotm'
    [void][IO.Directory]::CreateDirectory($target)
    $marker = Join-Path $target 'original.txt'
    [IO.File]::WriteAllText($marker, 'original must survive')
    $bad = Join-Path $root 'broken.dotm'
    [IO.File]::WriteAllText($bad, 'not a zip')
    Expect-Failure { Expand-RAToolsDotmToDirectory $bad $target -SafeRoot $root }
    Assert-True (Test-Path -LiteralPath $marker) 'Broken package must not remove the original directory'
    Assert-True ([IO.File]::ReadAllText($marker) -eq 'original must survive') 'Original bytes must remain'
    Write-Host 'PASS broken package preserves original'

    $incomplete = Join-Path $root 'incomplete.dotm'
    New-FixturePackage $incomplete $false
    Expect-Failure { Expand-RAToolsDotmToDirectory $incomplete $target -SafeRoot $root }
    Assert-True (Test-Path -LiteralPath $marker) 'Missing required part must preserve original'
    $valid = Join-Path $root 'valid.dotm'
    New-FixturePackage $valid
    Expect-Failure { Expand-RAToolsDotmToDirectory $valid $root -SafeRoot $root }
    Assert-True (Test-Path -LiteralPath $marker) 'Safe root itself must never be replaced'
    $inside = Join-Path $target 'inside.dotm'
    [IO.File]::Copy($valid, $inside)
    Expect-Failure { Expand-RAToolsDotmToDirectory $inside $target -SafeRoot $root }
    Assert-True (Test-Path -LiteralPath $inside) 'Input inside target must not be moved'
    Write-Host 'PASS incomplete package and unsafe targets preserve originals'

    # Inject a failure only in the final staging-directory rename.
    & $module {
        function script:Move-Item {
            param([string]$LiteralPath,[string]$Destination)
            if ((Split-Path -Leaf $LiteralPath) -like '.ratools-stage-*') { throw 'Injected publication failure' }
            Microsoft.PowerShell.Management\Move-Item -LiteralPath $LiteralPath -Destination $Destination
        }
    }
    Expect-Failure { Expand-RAToolsDotmToDirectory $valid $target -SafeRoot $root }
    Assert-True (Test-Path -LiteralPath $marker) 'Publication failure must restore the original target'
    & $module {
        function script:Move-Item {
            param([string]$LiteralPath,[string]$Destination)
            Microsoft.PowerShell.Management\Move-Item -LiteralPath $LiteralPath -Destination $Destination
        }
    }
    Expand-RAToolsDotmToDirectory $valid $target -SafeRoot $root | Out-Null
    Assert-True (Test-Path -LiteralPath (Join-Path $target 'word\vbaProject.bin')) 'Valid package must publish'
    Assert-True (-not (Test-Path -LiteralPath $marker)) 'Old content must leave the active directory'
    $backups = @(Get-ChildItem -LiteralPath (Join-Path $root 'backups') -Directory)
    Assert-True (@($backups | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'original.txt') }).Count -ge 1) 'Original directory must remain recoverable in backups'
    Write-Host 'PASS rollback and recoverable successful replacement'
    $output = Join-Path $root 'release.dotm'
    [IO.File]::WriteAllText($output, 'previous release')
    & $module {
        function script:Move-Item {
            param([string]$LiteralPath,[string]$Destination)
            if ((Split-Path -Leaf $LiteralPath) -like '.ratools-file-*') { throw 'Injected file publication failure' }
            Microsoft.PowerShell.Management\Move-Item -LiteralPath $LiteralPath -Destination $Destination -ErrorAction Stop
        }
    }
    Expect-Failure { Publish-RAToolsPackageFile $valid $output }
    Assert-True ([IO.File]::ReadAllText($output) -eq 'previous release') 'File publication failure must restore previous output'
    & $module {
        function script:Move-Item {
            param([string]$LiteralPath,[string]$Destination)
            Microsoft.PowerShell.Management\Move-Item -LiteralPath $LiteralPath -Destination $Destination -ErrorAction Stop
        }
    }
    Publish-RAToolsPackageFile $valid $output | Out-Null
    Assert-True ((Get-FileHash $valid).Hash -eq (Get-FileHash $output).Hash) 'Published output must equal source package'
    $oldOutputs = @(Get-ChildItem -LiteralPath (Join-Path $root 'backups') -File -Filter 'release.dotm.*.bak')
    Assert-True ($oldOutputs.Count -eq 1 -and [IO.File]::ReadAllText($oldOutputs[0].FullName) -eq 'previous release') 'Previous output must remain backed up'
    $before = (Get-FileHash $output).Hash
    $locked = [IO.File]::Open((Join-Path $target 'word\vbaProject.bin'), 'Open', 'Read', 'None')
    try { Expect-Failure { New-RAToolsDotmFromDirectory $target $output } } finally { $locked.Dispose() }
    Assert-True ((Get-FileHash $output).Hash -eq $before) 'Packing failure must preserve existing output'
    Expect-Failure { New-RAToolsDotmFromDirectory $target (Join-Path $target 'self.dotm') }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $target 'self.dotm'))) 'Output inside input must be rejected'
    Write-Host 'PASS output rollback, backup and failed packing preservation'
    $nestedOutput = Join-Path $target 'must-not-create\out.dotm'
    Expect-Failure { New-RAToolsDotmFromDirectory $target $nestedOutput }
    Assert-True (-not (Test-Path -LiteralPath (Split-Path -Parent $nestedOutput))) 'Reject output inside source before creating directories'
    & $module {
        function script:Get-ChildItem {
            [CmdletBinding()]
            param([string]$LiteralPath,[switch]$Recurse,[switch]$Force,[switch]$File)
            Write-Error 'Injected nonterminating enumeration failure'
        }
    }
    $priorPreference = $ErrorActionPreference
    $caught = $false
    try {
        $ErrorActionPreference = 'Continue'
        try { New-RAToolsDotmFromDirectory $target $output | Out-Null } catch { $caught = $true }
    } finally {
        $ErrorActionPreference = $priorPreference
        & $module {
            function script:Get-ChildItem {
                [CmdletBinding()]
                param([string]$LiteralPath,[switch]$Recurse,[switch]$Force,[switch]$File)
                Microsoft.PowerShell.Management\Get-ChildItem @PSBoundParameters
            }
        }
    }
    Assert-True $caught 'Enumeration errors must terminate even when caller uses Continue'
    Assert-True ((Get-FileHash $output).Hash -eq $before) 'Enumeration failure must keep previous package bytes'
    Write-Host 'PASS fail-closed enumeration and no side effects before rejection'
    $built = Join-Path $root 'dist\integrated.dotm'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'scripts\Build-RAToolsDotm.ps1') -RepoRoot $root -SkipVbaImport -OutputPath $built
    Assert-True ($LASTEXITCODE -eq 0) 'Build entry point must complete using safe publication'
    Assert-True (Test-Path -LiteralPath $built) 'Integrated build must produce its output'
    $archive = [IO.Compression.ZipFile]::OpenRead($built)
    try { Assert-True ($null -ne $archive.GetEntry('word/vbaProject.bin')) 'Built output must contain the VBA package part' }
    finally { $archive.Dispose() }
    Write-Host 'PASS integrated package build and directory synchronization'

    # Controlled junction points only to the fixture's own target; reviewed before recycling.
    $link = Join-Path $root 'linked-dotm'
    [void](New-Item -ItemType Junction -Path $link -Target $target)
    Expect-Failure { Expand-RAToolsDotmToDirectory $valid $link -SafeRoot $root }
    Assert-True (Test-Path -LiteralPath (Join-Path $target 'word\vbaProject.bin')) 'Linked target must be untouched'
    Expect-Failure { Move-RAToolsToRecycleBin $link -SafeRoot $root }
    Assert-True (Test-Path -LiteralPath $link) 'Linked paths must require review instead of automatic cleanup'
    Expect-Failure { New-RAToolsDotmFromDirectory $target (Join-Path $link 'linked-new\out.dotm') }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $target 'linked-new'))) 'Link rejection must precede creating output parent'
    $priorPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'SilentlyContinue'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'scripts\Build-RAToolsDotm.ps1') -RepoRoot $root -SkipVbaImport -NoSyncDotmDirectory -OutputPath (Join-Path $link 'entry-new\out.dotm') 2>$null
    } finally { $ErrorActionPreference = $priorPreference }
    Assert-True ($LASTEXITCODE -ne 0) 'Build entry must reject linked output'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $target 'entry-new'))) 'Build entry must not create directories before validating output'
    Write-Host 'PASS linked paths refused'
} finally {
    $resolved = [IO.Path]::GetFullPath($root)
    $allowed = [IO.Path]::GetFullPath((Join-Path $repo 'backups')).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected fixture cleanup path' }
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($resolved,'OnlyErrorDialogs','SendToRecycleBin')
}
Write-Host 'PASS BuildSafety.Tests'
