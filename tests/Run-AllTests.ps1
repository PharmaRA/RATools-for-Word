[CmdletBinding()]
param(
    # 跳过依赖 Word COM 的测试（CI 或未安装 Word 的环境使用）
    [switch]$SkipCom
)

$ErrorActionPreference = "Stop"

# 测试清单。RequiresCom 标记该脚本是否需要 Word COM 自动化。
$suites = @(
    @{ Name = "BuildRATools";            File = "BuildRATools.Tests.ps1";            RequiresCom = $false },
    @{ Name = "VersionConsistency";      File = "VersionConsistency.Tests.ps1";      RequiresCom = $false },
    @{ Name = "NormalizeScientificTerms"; File = "NormalizeScientificTerms.Tests.ps1"; RequiresCom = $true },
    @{ Name = "BatchDetectHighlights";   File = "BatchDetectHighlights.Tests.ps1";   RequiresCom = $true },
    @{ Name = "QuickToolbar";            File = "QuickToolbar.Tests.ps1";            RequiresCom = $true }
)

$results = @()
$failed = 0

foreach ($suite in $suites) {
    if ($SkipCom -and $suite.RequiresCom) {
        $results += [pscustomobject]@{ Suite = $suite.Name; Result = "SKIP (COM)" }
        continue
    }

    $path = Join-Path $PSScriptRoot $suite.File
    if (-not (Test-Path -LiteralPath $path)) {
        $results += [pscustomobject]@{ Suite = $suite.Name; Result = "MISSING" }
        $failed++
        continue
    }

    Write-Host ""
    Write-Host "===== $($suite.Name) =====" -ForegroundColor Cyan

    # 子进程运行，避免测试脚本之间的模块/全局状态串扰
    & powershell -NoProfile -ExecutionPolicy Bypass -File $path
    if ($LASTEXITCODE -eq 0) {
        $results += [pscustomobject]@{ Suite = $suite.Name; Result = "PASS" }
    }
    else {
        $results += [pscustomobject]@{ Suite = $suite.Name; Result = "FAIL" }
        $failed++
    }
}

Write-Host ""
Write-Host "===== Summary =====" -ForegroundColor Cyan
foreach ($r in $results) {
    $color = switch -Wildcard ($r.Result) {
        "PASS"   { "Green" }
        "SKIP*"  { "Yellow" }
        default  { "Red" }
    }
    Write-Host ("{0,-28} {1}" -f $r.Suite, $r.Result) -ForegroundColor $color
}

if ($failed -gt 0) {
    Write-Host ""
    Write-Host "FAIL Run-AllTests ($failed suite(s) failed)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "PASS Run-AllTests" -ForegroundColor Green
exit 0
