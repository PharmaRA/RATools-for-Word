[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Path,
    [int]$Tail = 0
)

$ErrorActionPreference = "Stop"
$gbk = [System.Text.Encoding]::GetEncoding(936)
$text = $gbk.GetString([System.IO.File]::ReadAllBytes($Path))
$lines = $text -split "`r`n"

if ($Tail -gt 0 -and $lines.Count -gt $Tail) {
    $start = $lines.Count - $Tail
} else {
    $start = 0
}

for ($i = $start; $i -lt $lines.Count; $i++) {
    "{0,4}: {1}" -f ($i + 1), $lines[$i]
}
