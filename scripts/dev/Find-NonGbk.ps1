<#
.SYNOPSIS
报告 UTF-8 文本文件中无法用 GBK(936) 表示的字符。
用于在写回 VBA 源码前定位不可编码字符（诊断辅助脚本）。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path
)

$ErrorActionPreference = "Stop"

$text = [System.IO.File]::ReadAllText(
    (Resolve-Path -LiteralPath $Path).Path,
    (New-Object System.Text.UTF8Encoding($false)))

$gbk = [System.Text.Encoding]::GetEncoding(936)
$lineNumber = 0

foreach ($line in ($text -split "`n")) {
    $lineNumber++
    for ($i = 0; $i -lt $line.Length; $i++) {
        $ch = $line[$i]
        if ([int]$ch -lt 128) { continue }
        $roundTrip = $gbk.GetString($gbk.GetBytes([string]$ch))
        if ($roundTrip -ne [string]$ch) {
            $code = [int]$ch
            Write-Host ("line={0} col={1} char=<{2}> U+{3:X4}" -f $lineNumber, $i, $ch, $code)
        }
    }
}

Write-Host "scan complete"
