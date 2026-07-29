[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [ValidateSet("gbk", "utf8")]
    [string]$As = "gbk"
)

$ErrorActionPreference = "Stop"

$gbk = [System.Text.Encoding]::GetEncoding(936)
$readEncoding = if ($As -eq "gbk") { $gbk } else { [System.Text.Encoding]::UTF8 }
$text = [System.IO.File]::ReadAllText($Path, $readEncoding)

$roundTrip = $gbk.GetString($gbk.GetBytes($text))

Write-Host "file=$Path readAs=$As chars=$($text.Length) roundTripChars=$($roundTrip.Length) equal=$($roundTrip -eq $text)"

if ($roundTrip -eq $text) { return }

$limit = [Math]::Min($text.Length, $roundTrip.Length)
$reported = 0
for ($i = 0; $i -lt $limit -and $reported -lt 20; $i++) {
    if ($text[$i] -ne $roundTrip[$i]) {
        $before = [int][char]$text[$i]
        $after = [int][char]$roundTrip[$i]
        $lineNo = ($text.Substring(0, $i) -split "`n").Count
        $ctxStart = [Math]::Max(0, $i - 20)
        $ctx = $text.Substring($ctxStart, [Math]::Min(40, $text.Length - $ctxStart)) -replace "`r?`n", " / "
        Write-Host ("line={0} idx={1} U+{2:X4} -> U+{3:X4}  ctx=<{4}>" -f $lineNo, $i, $before, $after, $ctx)
        $reported++
    }
}
