$ErrorActionPreference = "Stop"

# Ribbon 回调完整性校验（纯静态，无 COM）：
# customUI14.xml 中引用的每个 onAction/onLoad 回调，必须在 modules/*.bas 中
# 存在同名 Public Sub。防止拆分/改名导致按钮静默失效
# （Word 对缺失回调不报错，只是点击无反应）。

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$ribbonPath = Join-Path $repoRoot "dotm\customUI\customUI14.xml"
$modulesDir = Join-Path $repoRoot "modules"

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

Assert-True (Test-Path -LiteralPath $ribbonPath) "Ribbon XML missing: $ribbonPath"

$ribbonText = Get-Content -LiteralPath $ribbonPath -Raw

# 1. 收集 XML 里引用的全部回调名
$callbacks = New-Object System.Collections.Generic.HashSet[string]
foreach ($m in [regex]::Matches($ribbonText, 'on(?:Action|Load)="([^"]+)"')) {
    [void]$callbacks.Add($m.Groups[1].Value)
}
# 当前 XML 为 17 个去重 onAction + 1 个 onLoad；阈值防解析退化
Assert-True ($callbacks.Count -ge 15) "Suspiciously few callbacks found: $($callbacks.Count)"
Write-Host "Ribbon callbacks referenced: $($callbacks.Count)"

# 2. 收集全部标准模块的 Public Sub 名（GBK 读取）
$gbk = [System.Text.Encoding]::GetEncoding(936)
$publicSubs = New-Object System.Collections.Generic.HashSet[string]
foreach ($file in Get-ChildItem -LiteralPath $modulesDir -Filter "*.bas" -File) {
    $source = [System.IO.File]::ReadAllText($file.FullName, $gbk)
    foreach ($m in [regex]::Matches($source, '(?m)^(?:Public\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(')) {
        # 模块级默认可见性为 Public；显式 Private 的已被上面的正则排除（不含 Private 前缀才匹配）
        [void]$publicSubs.Add($m.Groups[1].Value)
    }
}
Write-Host "Public subs found: $($publicSubs.Count)"

# 3. 比对
$missing = @($callbacks | Where-Object { -not $publicSubs.Contains($_) })
if ($missing.Count -gt 0) {
    $missing | ForEach-Object { Write-Host "MISSING CALLBACK: $_" -ForegroundColor Red }
    throw "$($missing.Count) Ribbon callback(s) have no matching Public Sub"
}

Write-Host "PASS RibbonCallbacks.Tests"
