<#
.SYNOPSIS
Report raw byte patterns for non-ASCII runs in a source file, to verify GBK(936) storage.

.DESCRIPTION
GBK Chinese is 2 bytes per character with lead byte 0x81-0xFE.
UTF-8 Chinese is 3 bytes per character with lead byte 0xE0-0xEF.
This prints the first non-ASCII run found at or after a given anchor line so the
two encodings can be told apart by eye without depending on console code pages.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [int]$MaxRuns = 6
)

$ErrorActionPreference = "Stop"

$bytes = [System.IO.File]::ReadAllBytes($Path)
Write-Output ("file={0}" -f $Path)
Write-Output ("totalBytes={0}" -f $bytes.Length)

# Walk the file, collecting maximal runs of bytes >= 0x80, reporting a few samples.
$runs = New-Object System.Collections.Generic.List[object]
$i = 0
$line = 1
while ($i -lt $bytes.Length) {
    if ($bytes[$i] -eq 0x0A) { $line++ }
    if ($bytes[$i] -ge 0x80) {
        $start = $i
        while ($i -lt $bytes.Length -and $bytes[$i] -ge 0x80) { $i++ }
        $runs.Add([pscustomobject]@{
            Line   = $line
            Length = $i - $start
            Hex    = (($bytes[$start..($i - 1)] | ForEach-Object { $_.ToString("X2") }) -join " ")
        })
    }
    else {
        $i++
    }
}

Write-Output ("nonAsciiRuns={0}" -f $runs.Count)

# Classify: count runs whose length is divisible by 3 with 0xE_ leads (UTF-8 shape).
$utf8Shaped = 0
$gbkShaped = 0
foreach ($run in $runs) {
    $leads = ($run.Hex -split " ")
    $firstLead = [Convert]::ToInt32($leads[0], 16)
    if ($firstLead -ge 0xE0 -and $firstLead -le 0xEF -and ($run.Length % 3) -eq 0) {
        $utf8Shaped++
    }
    elseif ($firstLead -ge 0x81 -and $firstLead -le 0xFE -and ($run.Length % 2) -eq 0) {
        $gbkShaped++
    }
}
Write-Output ("utf8ShapedRuns={0}" -f $utf8Shaped)
Write-Output ("gbkShapedRuns={0}" -f $gbkShaped)

Write-Output "--- first runs (line / byteLen / hex) ---"
foreach ($run in ($runs | Select-Object -First $MaxRuns)) {
    Write-Output ("line={0,4}  len={1,3}  {2}" -f $run.Line, $run.Length, $run.Hex)
}

Write-Output "--- last runs (line / byteLen / hex) ---"
foreach ($run in ($runs | Select-Object -Last $MaxRuns)) {
    Write-Output ("line={0,4}  len={1,3}  {2}" -f $run.Line, $run.Length, $run.Hex)
}
