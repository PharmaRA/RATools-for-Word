<#
.SYNOPSIS
Append a UTF-8 fragment file to a GBK(936) VBA source file.

.DESCRIPTION
Repository .bas/.cls/.frm files are stored as GBK(936); see
docs/vba-source-encoding.md. New code containing Chinese must never be passed
through a shell command line or heredoc: the terminal may re-encode it and drop
line breaks. Instead write the new text to a UTF-8 file and let this script do
the UTF-8 -> GBK conversion via .NET.

NOTE: this script deliberately contains ASCII comments only. PowerShell 5.1
reads a .ps1 without BOM using the ANSI code page, so UTF-8 Chinese comments
would be mis-decoded and can swallow the following line break, silently
commenting out the next statement.

.PARAMETER Path
Target VBA source file (stored as GBK).

.PARAMETER FragmentPath
UTF-8 file holding the text to append.

.PARAMETER AssertAbsent
Assert the target does not already contain these strings, to prevent a
duplicate append. May be given more than once.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [Parameter(Mandatory)]
    [string]$FragmentPath,

    [string[]]$AssertAbsent = @()
)

$ErrorActionPreference = "Stop"

foreach ($required in @($Path, $FragmentPath)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "File not found: $required"
    }
}

$gbk = [System.Text.Encoding]::GetEncoding(936)
$existing = [System.IO.File]::ReadAllText($Path, $gbk)

foreach ($needle in $AssertAbsent) {
    if ($existing.Contains($needle)) {
        throw "Refusing to append: target already contains <$needle>"
    }
}

# Read the fragment as UTF-8 (.NET honours a BOM when present), normalise to
# CRLF because VBA source is always CRLF, then append.
$fragment = [System.IO.File]::ReadAllText($FragmentPath, [System.Text.Encoding]::UTF8)
$fragment = $fragment -replace "`r`n", "`n" -replace "`n", "`r`n"

if (-not $existing.EndsWith("`r`n")) { $existing += "`r`n" }
$combined = $existing + $fragment
if (-not $combined.EndsWith("`r`n")) { $combined += "`r`n" }

# Characters GBK cannot represent become '?' and would silently corrupt the
# source, so verify a lossless round trip before writing.
$roundTrip = $gbk.GetString($gbk.GetBytes($combined))
if ($roundTrip -ne $combined) {
    throw "Content is not representable in GBK(936); refusing to write $Path"
}

[System.IO.File]::WriteAllText($Path, $combined, $gbk)
Write-Host "OK appended $([System.IO.Path]::GetFileName($FragmentPath)) -> $Path"
