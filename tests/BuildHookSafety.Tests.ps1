$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
$root=Join-Path (Join-Path $repo 'backups') ('RATools_HookSafety_'+[guid]::NewGuid().ToString('N'))
foreach($dir in @('modules','class_modules','userforms','dotm','scripts','empty-hooks')){[void][IO.Directory]::CreateDirectory((Join-Path $root $dir))}
$utf8=[Text.UTF8Encoding]::new($false)
[IO.File]::WriteAllText((Join-Path $root 'modules\Engine_UpdateChecker.bas'),'baseline',$utf8)
[IO.File]::WriteAllText((Join-Path $root 'dotm\marker.txt'),'baseline',$utf8)
[IO.File]::WriteAllText((Join-Path $root 'CHANGELOG.md'),"# v1.0.0`n",$utf8)
[IO.File]::WriteAllText((Join-Path $root 'scripts\Build-RAToolsDotm.ps1'),@'
$root=Split-Path -Parent $PSScriptRoot
[IO.File]::WriteAllText((Join-Path $root 'built.txt'),[IO.File]::ReadAllText((Join-Path $root 'CHANGELOG.md')))
'@,$utf8)
function Run-Hook {
    $prior=$ErrorActionPreference
    try {
        $ErrorActionPreference='Continue'
        $output=& 'C:\Program Files\Git\bin\bash.exe' '.git/hooks/pre-commit' 2>&1
        return [pscustomobject]@{Code=$LASTEXITCODE;Output=$output}
    } finally {$ErrorActionPreference=$prior}
}
try {
    & git -C $root init -q
    & git -C $root add -- .
    & git -C $root -c user.name=Fiveonine -c user.email=57358895+Fiveo9@users.noreply.github.com -c core.hooksPath=empty-hooks commit -qm 'test baseline'
    if($LASTEXITCODE -ne 0){throw 'Fixture baseline failed'}
    & (Join-Path $repo 'scripts\Install-RAToolsBuildHook.ps1') -RepoRoot $root
    Push-Location $root
    try {
        # The index has v1.1.0; unstaged edits advance the working version again.
        [IO.File]::WriteAllText((Join-Path $root 'CHANGELOG.md'),"# v1.1.0`n",$utf8)
        & git add -- CHANGELOG.md
        [IO.File]::WriteAllText((Join-Path $root 'CHANGELOG.md'),"# v1.2.0`n",$utf8)
        $result=Run-Hook
        if($result.Code -eq 0){throw 'Partial CHANGELOG staging must reject the hook before build'}
        if(Test-Path -LiteralPath (Join-Path $root 'built.txt')){throw 'Rejected input must not run the builder'}
        Write-Host 'PASS partially staged CHANGELOG rejects before building'
        & git add -- CHANGELOG.md
        $result=Run-Hook
        if($result.Code -ne 0){throw "Fully staged input should build: $($result.Output)"}
        if([IO.File]::ReadAllText((Join-Path $root 'built.txt')) -ne "# v1.2.0`n"){throw 'Builder must see the staged version'}
        Write-Host 'PASS fully staged version builds'
        Move-RAToolsToRecycleBin (Join-Path $root 'built.txt') -SafeRoot $root
        [IO.File]::WriteAllText((Join-Path $root 'modules\Engine_UpdateChecker.bas'),'unstaged source',$utf8)
        $result=Run-Hook
        if($result.Code -eq 0 -or (Test-Path -LiteralPath (Join-Path $root 'built.txt'))){throw 'Unstaged VBA must still reject before build'}
        Write-Host 'PASS original VBA staging guard remains effective'
    } finally {Pop-Location}
} finally {
    Move-RAToolsToRecycleBin $root -SafeRoot (Join-Path $repo 'backups')
}
Write-Host 'PASS BuildHookSafety.Tests'
