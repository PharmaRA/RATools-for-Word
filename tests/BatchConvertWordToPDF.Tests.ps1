#requires -version 5.1
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

# Macro_BatchConvertWordToPDF 转换选项与域锁定机制测试。

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $repoRoot "modules\Macro_BatchConvertWordToPDF.bas"
$macroRegistryPath = Join-Path $repoRoot "modules\UI_MacroRegistry.bas"

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Read-VbaSource {
    param([string]$Path)
    $vbaEncoding = [System.Text.Encoding]::GetEncoding(936)
    return $vbaEncoding.GetString([System.IO.File]::ReadAllBytes($Path))
}

Write-Output "Running BatchConvertWordToPDF source structure checks"
$moduleText = Read-VbaSource -Path $modulePath
$macroRegistryText = Read-VbaSource -Path $macroRegistryPath

Assert-True ($moduleText.Contains('Private Const CONVERT_OPT_WYSIWYG As String = "1"')) `
    "Module should define CONVERT_OPT_WYSIWYG as option 1."
Assert-True ($moduleText.Contains('Private Const CONVERT_OPT_PAGE_ONLY As String = "2"')) `
    "Module should define CONVERT_OPT_PAGE_ONLY as option 2."
Assert-True ($moduleText.Contains('Private Const CONVERT_OPT_FULL_TOC As String = "3"')) `
    "Module should define CONVERT_OPT_FULL_TOC as option 3."
Assert-True ($moduleText.Contains("Public Sub SetAllFieldsLocked(doc As Document, ByVal lockState As Boolean)")) `
    "Module should implement SetAllFieldsLocked helper."
Assert-True ($moduleText.Contains("Private Function ChooseConvertOption(ByVal dialogTitle As String) As String")) `
    "Module should implement ChooseConvertOption dialog."
Assert-True ($moduleText.Contains("SetAllFieldsLocked doc, True")) `
    "Module should lock fields before WYSIWYG export."
Assert-True ($moduleText.Contains("SetAllFieldsLocked doc, False")) `
    "Module should unlock fields after WYSIWYG export."
Assert-True ($macroRegistryText.Contains("BatchConvertWordToPDF")) `
    "Macro registry should register BatchConvertWordToPDF."

Write-Output "Source checks PASS"

$word = $null
$doc = $null
try {
    $word = Start-RAToolsWordSession
    $doc = $word.Documents.Add()
    [void]$doc.VBProject.VBComponents.Import((Join-Path $repoRoot "modules\Core_Files.bas"))
    [void]$doc.VBProject.VBComponents.Import((Join-Path $repoRoot "modules\Core_Session.bas"))
    [void]$doc.VBProject.VBComponents.Import((Join-Path $repoRoot "modules\Core_UI.bas"))
    [void]$doc.VBProject.VBComponents.Import($modulePath)

    # 1. 主文档正文域
    $mainFld = $doc.Fields.Add($doc.Range(0, 0), 33, "PAGE")

    # 2. 页眉域
    $section = $doc.Sections.Item(1)
    $header = $section.Headers.Item(1)
    $headerFld = $header.Range.Fields.Add($header.Range, 33, "PAGE")

    # 3. 形状/文本框域
    $shape = $doc.Shapes.AddTextbox(1, 50, 50, 150, 80)
    $shape.TextFrame.TextRange.Text = "Probe "
    $shapeFld = $shape.TextFrame.TextRange.Fields.Add($shape.TextFrame.TextRange, 33, "PAGE")

    # 测试一键锁定全部域
    [void]$word.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $word, @("SetAllFieldsLocked", $doc, $true))
    Assert-True ($mainFld.Locked -eq $true) "Main body field should be locked."
    Assert-True ($headerFld.Locked -eq $true) "Header field should be locked."
    Assert-True ($shapeFld.Locked -eq $true) "Shape textbox field should be locked."
    Write-Host "SetAllFieldsLocked(True): all story ranges and shapes locked OK"

    # 测试一键解锁全部域
    [void]$word.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $word, @("SetAllFieldsLocked", $doc, $false))
    Assert-True ($mainFld.Locked -eq $false) "Main body field should be unlocked."
    Assert-True ($headerFld.Locked -eq $false) "Header field should be unlocked."
    Assert-True ($shapeFld.Locked -eq $false) "Shape textbox field should be unlocked."
    Write-Host "SetAllFieldsLocked(False): all story ranges and shapes unlocked OK"
}
finally {
    if ($null -ne $doc) { $doc.Saved = $true; $doc.Close(0) }
    Stop-RAToolsWordSession -Word $word
}

Write-Host "PASS BatchConvertWordToPDF.Tests"
