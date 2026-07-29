[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Path,
    [Parameter(Mandatory)] [string]$FragmentPath
)

$ErrorActionPreference = "Stop"
$gbk = [System.Text.Encoding]::GetEncoding(936)

$existing = [System.IO.File]::ReadAllText($Path, $gbk)
$fragment = [System.IO.File]::ReadAllText($FragmentPath, [System.Text.Encoding]::UTF8)

Write-Host "existingChars=$($existing.Length) fragmentChars=$($fragment.Length)"
Write-Host ("fragment first char = U+{0:X4}" -f [int][char]$fragment[0])

$fragment = $fragment -replace "`r`n", "`n" -replace "`n", "`r`n"
if (-not $existing.EndsWith("`r`n")) { $existing += "`r`n" }
$combined = $existing + $fragment

$roundTrip = $gbk.GetString($gbk.GetBytes($combined))
Write-Host "combinedChars=$($combined.Length) roundTrip=$($roundTrip.Length) equal=$($roundTrip -eq $combined)"

$limit = [Math]::Min($combined.Length, $roundTrip.Length)
$reported = 0
for ($i = 0; $i -lt $limit -and $reported -lt 10; $i++) {
    if ($combined[$i] -ne $roundTrip[$i]) {
        $lineNo = ($combined.Substring(0, $i) -split "`n").Count
        Write-Host ("line={0} idx={1} U+{2:X4} -> U+{3:X4}" -f $lineNo, $i, [int][char]$combined[$i], [int][char]$roundTrip[$i])
        $reported++
    }
}
