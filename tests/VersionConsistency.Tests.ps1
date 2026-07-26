$ErrorActionPreference = "Stop"

# 校验版本号三处一致：CHANGELOG.md 顶部标题、Mod_UpdateChecker.bas 的 APP_VERSION、README.md 的“当前发布版本”。
# 全程纯文本读取，无 Word COM 依赖，可在 CI 运行。

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$modulePath = Join-Path $repoRoot "scripts\RATools.Build.psm1"

Import-Module $modulePath -Force

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Run-Test {
    param(
        [string]$Name,
        [scriptblock]$Body
    )

    Write-Host "Running $Name"
    & $Body
}

$updateCheckerPath = Join-Path $repoRoot "modules\Mod_UpdateChecker.bas"
$readmePath = Join-Path $repoRoot "README.md"

$changelogVersion = Get-RAToolsLatestChangelogVersion -RepoRoot $repoRoot

Run-Test "Changelog top version is a vX.Y.Z tag" {
    Assert-True ($changelogVersion -match '^v\d+\.\d+\.\d+$') `
        "Unexpected changelog version format: $changelogVersion"
}

Run-Test "APP_VERSION matches changelog version" {
    $source = Get-Content -LiteralPath $updateCheckerPath -Raw
    $match = [Regex]::Match($source, 'Private\s+Const\s+APP_VERSION\s+As\s+String\s*=\s*"([^"]+)"')
    Assert-True $match.Success "APP_VERSION constant not found in Mod_UpdateChecker.bas"
    Assert-True ($match.Groups[1].Value -eq $changelogVersion) `
        "APP_VERSION <$($match.Groups[1].Value)> does not match changelog <$changelogVersion>"
}

Run-Test "README current release version matches changelog version" {
    # README.md 为 UTF-8 无 BOM，PowerShell 5.1 需显式指定编码
    $readme = [System.IO.File]::ReadAllText($readmePath, [System.Text.Encoding]::UTF8)
    $match = [Regex]::Match($readme, '当前发布版本：`(v\d+\.\d+\.\d+)`')
    Assert-True $match.Success "README current-release line not found (expected 当前发布版本：`vX.Y.Z`)"
    Assert-True ($match.Groups[1].Value -eq $changelogVersion) `
        "README version <$($match.Groups[1].Value)> does not match changelog <$changelogVersion>"
}

Write-Host "PASS VersionConsistency.Tests"
