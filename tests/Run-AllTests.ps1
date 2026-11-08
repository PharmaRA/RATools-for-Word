[CmdletBinding()]
param(
    # 跳过依赖 Word COM 的测试（CI 或未安装 Word 的环境使用）
    [switch]$SkipCom
)

$ErrorActionPreference = "Stop"

# 测试清单。RequiresCom 标记该脚本是否需要 Word COM 自动化。
$suites = @(
    @{ Name = "BuildRATools";            File = "BuildRATools.Tests.ps1";            RequiresCom = $false },
    @{ Name = "BuildSafety";             File = "BuildSafety.Tests.ps1";             RequiresCom = $false },
    @{ Name = "VersionConsistency";      File = "VersionConsistency.Tests.ps1";      RequiresCom = $false },
    @{ Name = "RibbonCallbacks";         File = "RibbonCallbacks.Tests.ps1";         RequiresCom = $false },
    @{ Name = "QuickToolbarWindow";      File = "QuickToolbarWindow.Tests.ps1";      RequiresCom = $false },
    @{ Name = "ScriptEncoding";          File = "ScriptEncoding.Tests.ps1";          RequiresCom = $false },
    @{ Name = "FormatRevisionPath"; File = "FormatRevisionPath.Tests.ps1"; RequiresCom = $true },
    @{ Name = "AcceptFormatRevisions"; File = "AcceptFormatRevisions.Tests.ps1"; RequiresCom = $true },
    @{ Name = "SourceCompile";           File = "SourceCompile.Tests.ps1";           RequiresCom = $true },
    @{ Name = "StyleNameMapping";        File = "StyleNameMapping.Tests.ps1";        RequiresCom = $true },
    @{ Name = "UnusedStylesSafety";       File = "UnusedStylesSafety.Tests.ps1";       RequiresCom = $true },
    @{ Name = "UpdateChecker";           File = "UpdateChecker.Tests.ps1";           RequiresCom = $true },
    @{ Name = "BatchRenameFiles";        File = "BatchRenameFiles.Tests.ps1";        RequiresCom = $true },
    @{ Name = "RenamePathSafety";        File = "RenamePathSafety.Tests.ps1";        RequiresCom = $true },
    @{ Name = "NormalizeScientificTerms"; File = "NormalizeScientificTerms.Tests.ps1"; RequiresCom = $true },
    @{ Name = "BatchDetectHighlights";   File = "BatchDetectHighlights.Tests.ps1";   RequiresCom = $true },
    @{ Name = "BatchConvertWordToPDF";   File = "BatchConvertWordToPDF.Tests.ps1";   RequiresCom = $true },
    @{ Name = "PdfOutputSafety";          File = "PdfOutputSafety.Tests.ps1";          RequiresCom = $true },
    @{ Name = "PdfFieldState";            File = "PdfFieldState.Tests.ps1";            RequiresCom = $true },
    @{ Name = "QuickToolbar";            File = "QuickToolbar.Tests.ps1";            RequiresCom = $true },
    @{ Name = "ReviewerFilter";          File = "ReviewerFilter.Tests.ps1";          RequiresCom = $true },
    @{ Name = "ReviewerStorySafety";     File = "ReviewerStorySafety.Tests.ps1";     RequiresCom = $true },
    @{ Name = "PasteRelinkSafety"; File = "PasteRelinkSafety.Tests.ps1"; RequiresCom = $true },
    @{ Name = "PasteRelinkCrossReference"; File = "PasteRelinkCrossReference.Tests.ps1"; RequiresCom = $true },
    @{ Name = "PasteRelinkCrossReferenceCases"; File = "PasteRelinkCrossReferenceCases.Tests.ps1"; RequiresCom = $true },
    @{ Name = "ReviewerFormRuntime";     File = "ReviewerFormRuntime.Tests.ps1";     RequiresCom = $true }
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
