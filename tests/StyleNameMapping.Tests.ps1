$ErrorActionPreference = "Stop"

# 校验 Mod_StyleNames 的中英映射表与 template/master-template-en.dotx 的
# 实际样式名一致：每个映射目标英文名必须存在于英文模板中。
# 通过解包 dotx 读取 word/styles.xml 的 w:name 清单（纯文件解析，但依赖
# Word COM 导入源码执行 GetStyleNameMappingData，故归类 COM 测试）。

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$templatePath = Join-Path $repoRoot "template\master-template-en.dotx"
$styleModulePath = Join-Path $repoRoot "modules\Mod_StyleNames.bas"

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

Assert-True (Test-Path -LiteralPath $templatePath) "English template missing: $templatePath"

# --- 1. 从 dotx 提取全部样式名 ---
Add-Type -AssemblyName System.IO.Compression.FileSystem
$templateStyles = New-Object System.Collections.Generic.HashSet[string]
$archive = [System.IO.Compression.ZipFile]::OpenRead($templatePath)
try {
    $entry = $archive.GetEntry("word/styles.xml")
    Assert-True ($null -ne $entry) "word/styles.xml missing in template"
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { $stylesXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
}
finally {
    $archive.Dispose()
}

[xml]$doc = $stylesXml
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
foreach ($nameNode in $doc.SelectNodes("//w:style/w:name", $ns)) {
    [void]$templateStyles.Add($nameNode.GetAttribute("w:val"))
}
# 别名（w:aliases）也计入
foreach ($aliasNode in $doc.SelectNodes("//w:style/w:aliases", $ns)) {
    foreach ($alias in $aliasNode.GetAttribute("w:val").Split(",")) {
        [void]$templateStyles.Add($alias.Trim())
    }
}

Assert-True ($templateStyles.Count -gt 10) "Template style list suspiciously small: $($templateStyles.Count)"
Write-Host "Template styles loaded: $($templateStyles.Count)"

# --- 2. 用 Word COM 执行 GetStyleNameMappingData 取回映射表 ---
$word = $null
$comDoc = $null
$mappingLines = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1
    $comDoc = $word.Documents.Add()
    [void]$comDoc.VBProject.VBComponents.Import($styleModulePath)

    $probe = $comDoc.VBProject.VBComponents.Add(1)
    $probe.Name = "StyleMappingProbe"
    $probe.CodeModule.AddFromString(@'
Public Function DumpStyleMapping() As String
    Dim data As Collection
    Dim i As Long
    Dim result As String

    Set data = GetStyleNameMappingData()
    For i = 1 To data.count
        result = result & data(i)(0) & "|" & data(i)(1) & vbLf
    Next i
    DumpStyleMapping = result
End Function
'@)

    $mappingLines = ([string]$word.Run("DumpStyleMapping")).Split("`n") | Where-Object { $_ -ne "" }
}
finally {
    if ($null -ne $comDoc) { $comDoc.Saved = $true; $comDoc.Close(0) }
    if ($null -ne $word) { $word.Quit() }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Assert-True ($mappingLines.Count -ge 40) "Mapping table too small: $($mappingLines.Count)"
Write-Host "Mapping entries: $($mappingLines.Count)"

# --- 3. 每个英文名必须在模板中存在 ---
$missing = @()
$seenChinese = New-Object System.Collections.Generic.HashSet[string]
foreach ($line in $mappingLines) {
    $parts = $line.Split("|")
    Assert-True ($parts.Count -eq 2) "Malformed mapping line: $line"
    $cn = $parts[0]; $en = $parts[1]

    Assert-True $seenChinese.Add($cn) "Duplicate Chinese key in mapping: $cn"
    Assert-True ($cn.EndsWith("-F")) "Chinese name should end with -F: $cn"
    Assert-True ($en.EndsWith("-F")) "English name should end with -F: $en"

    if (-not $templateStyles.Contains($en)) {
        $missing += "$cn -> $en"
    }
}

if ($missing.Count -gt 0) {
    $missing | ForEach-Object { Write-Host "MISSING IN TEMPLATE: $_" -ForegroundColor Red }
    throw "$($missing.Count) mapped English style name(s) not found in master-template-en.dotx"
}

Write-Host "PASS StyleNameMapping.Tests"
