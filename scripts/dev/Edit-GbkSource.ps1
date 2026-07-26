<#
.SYNOPSIS
对 GBK 编码的 VBA 源码做精确字符串替换，保持编码不变。

.DESCRIPTION
仓库内 .bas/.cls/.frm 以 GBK(936) 存盘（Word/VBE 在中文系统的导入约定）。
通用文本编辑工具往往按 UTF-8 写回导致中文损坏，因此源码修改统一走本脚本：
按 936 读入 -> 精确替换（可多组）-> 按 936 写回。
Old/New 若含中文，调用方脚本自身应为 UTF-8 BOM 编码，PowerShell 5.1 会正确解码。

.PARAMETER Path
目标文件路径。

.PARAMETER Old
待替换的原文（必须与文件解码后的文本精确匹配）。

.PARAMETER New
替换后的文本。

.PARAMETER ExpectCount
Old 在文件中出现的期望次数（默认 1）；不匹配则报错退出，防止误替换。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [Parameter(Mandatory)]
    [string]$Old,

    [Parameter(Mandatory)]
    [AllowEmptyString()]
    [string]$New,

    [int]$ExpectCount = 1
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    throw "File not found: $Path"
}

$gbk = [System.Text.Encoding]::GetEncoding(936)
$text = [System.IO.File]::ReadAllText($Path, $gbk)

# VBA 源码固定为 CRLF；调用方脚本里的多行 here-string 可能是裸 LF，统一归一化
$Old = $Old -replace "`r`n", "`n" -replace "`n", "`r`n"
$New = $New -replace "`r`n", "`n" -replace "`n", "`r`n"

$count = ([System.Text.RegularExpressions.Regex]::Matches(
    $text, [System.Text.RegularExpressions.Regex]::Escape($Old))).Count

if ($count -ne $ExpectCount) {
    throw "Occurrence mismatch in ${Path}: expected $ExpectCount, found $count for anchor <$Old>"
}

$text = $text.Replace($Old, $New)
[System.IO.File]::WriteAllText($Path, $text, $gbk)

Write-Host "OK $Path ($ExpectCount replacement(s))"
