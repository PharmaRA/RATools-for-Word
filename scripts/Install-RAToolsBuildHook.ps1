#requires -version 5.1

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $PSCommandPath
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $scriptRoot "..")).Path
}

$gitDir = (& git -C $RepoRoot rev-parse --git-dir) 2>$null
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($gitDir)) {
    throw "Could not find a git repository at $RepoRoot."
}

if (-not [System.IO.Path]::IsPathRooted($gitDir)) {
    $gitDir = Join-Path $RepoRoot $gitDir
}

$hooksDir = Join-Path $gitDir "hooks"
$hookPath = Join-Path $hooksDir "pre-commit"

if (-not (Test-Path -LiteralPath $hooksDir -PathType Container)) {
    New-Item -ItemType Directory -Path $hooksDir | Out-Null
}

if ($Uninstall) {
    if (Test-Path -LiteralPath $hookPath -PathType Leaf) {
        Remove-Item -LiteralPath $hookPath -Force
        Write-Host "Removed RATools pre-commit build hook: $hookPath"
    }
    else {
        Write-Host "No pre-commit hook found at: $hookPath"
    }

    return
}

$hookLines = @(
    '#!/bin/sh',
    'set -e',
    "",
    'repo_root="$(git rev-parse --show-toplevel)"',
    "",
    'if ! command -v powershell.exe >/dev/null 2>&1; then',
    '  echo "powershell.exe was not found; cannot run RATools local build." >&2',
    '  exit 1',
    'fi',
    "",
    'if ! git diff --quiet -- modules class_modules userforms; then',
    '  echo "Unstaged VBA source changes found. Stage source files before committing so dotm matches the commit." >&2',
    '  exit 1',
    'fi',
    "",
    'untracked_sources="$(git ls-files --others --exclude-standard -- modules class_modules userforms)"',
    'if [ -n "$untracked_sources" ]; then',
    '  echo "Untracked VBA source files found. Stage or remove them before committing:" >&2',
    '  echo "$untracked_sources" >&2',
    '  exit 1',
    'fi',
    "",
    'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$repo_root/scripts/Build-RAToolsDotm.ps1"',
    'git add "$repo_root/modules/Mod_UpdateChecker.bas"',
    'git add "$repo_root/dotm"'
)
$hookContent = ($hookLines -join "`n") + "`n"

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($hookPath, $hookContent, $utf8NoBom)

Write-Host "Installed RATools pre-commit build hook: $hookPath"
