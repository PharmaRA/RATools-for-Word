<#
.SYNOPSIS
Replace an exact text block inside a GBK(936) VBA source file, reading both the
anchor and the replacement from UTF-8 files.

.DESCRIPTION
Companion to Edit-GbkSource.ps1 (command-line Old/New) and Append-GbkSource.ps1
(append only). Use this one when the anchor or the replacement contains Chinese:
passing Chinese through a shell command line or a here-string in a BOM-less
.ps1 is unreliable, so both sides are read from UTF-8 files instead.

Reads the target as code page 936, replaces OldPath's content with NewPath's
content (exact match, occurrence count asserted), verifies the result still
round-trips through GBK, then writes it back as 936. See
docs/vba-source-encoding.md.

NOTE: this script deliberately contains ASCII comments only. PowerShell 5.1
reads a .ps1 without BOM using the ANSI code page, so UTF-8 Chinese comments
would be mis-decoded and can swallow the following line break.

.PARAMETER Path
Target VBA source file (stored as GBK).

.PARAMETER OldPath
UTF-8 file holding the exact text to replace.

.PARAMETER NewPath
UTF-8 file holding the replacement text.

.PARAMETER ExpectCount
Expected number of occurrences of the anchor (default 1); a mismatch aborts
without writing, to prevent an unintended replacement.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [Parameter(Mandatory)]
    [string]$OldPath,

    [Parameter(Mandatory)]
    [string]$NewPath,

    [int]$ExpectCount = 1
)

$ErrorActionPreference = "Stop"

foreach ($required in @($Path, $OldPath, $NewPath)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "File not found: $required"
    }
}

$gbk = [System.Text.Encoding]::GetEncoding(936)
$text = [System.IO.File]::ReadAllText($Path, $gbk)

function Read-Utf8Block {
    param([string]$BlockPath)

    # .NET honours a BOM when present; normalise to CRLF because VBA source is
    # always CRLF, and drop a single trailing newline so the fragment files can
    # end with one without affecting the match.
    $block = [System.IO.File]::ReadAllText($BlockPath, [System.Text.Encoding]::UTF8)
    $block = $block -replace "`r`n", "`n" -replace "`n", "`r`n"
    if ($block.EndsWith("`r`n")) { $block = $block.Substring(0, $block.Length - 2) }
    return $block
}

$old = Read-Utf8Block -BlockPath $OldPath
$new = Read-Utf8Block -BlockPath $NewPath

if ($old.Length -eq 0) {
    throw "Anchor file is empty: $OldPath"
}

$count = ([System.Text.RegularExpressions.Regex]::Matches(
    $text, [System.Text.RegularExpressions.Regex]::Escape($old))).Count

if ($count -ne $ExpectCount) {
    throw "Occurrence mismatch in ${Path}: expected $ExpectCount, found $count for anchor from $OldPath"
}

$updated = $text.Replace($old, $new)

# Characters GBK cannot represent would become '?' and silently corrupt the
# source, so verify a lossless round trip before writing.
$roundTrip = $gbk.GetString($gbk.GetBytes($updated))
if ($roundTrip -ne $updated) {
    throw "Content is not representable in GBK(936); refusing to write $Path"
}

[System.IO.File]::WriteAllText($Path, $updated, $gbk)
Write-Host "OK $Path ($ExpectCount replacement(s) from $([System.IO.Path]::GetFileName($NewPath)))"
