#requires -version 5.1
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

# Macro_PasteRelinkCrossReference 测试：
# 源码静态契约 + Word COM 真实「复制粘贴后重连交叉引用」验证。

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $repoRoot "modules\Macro_PasteRelinkCrossReference.bas"
$macroRegistryPath = Join-Path $repoRoot "modules\UI_MacroRegistry.bas"
$ribbonPath = Join-Path $repoRoot "dotm\customUI\customUI14.xml"

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Read-VbaSource {
    param([string]$Path)
    $vbaEncoding = [System.Text.Encoding]::GetEncoding(936)
    return $vbaEncoding.GetString([System.IO.File]::ReadAllBytes($Path))
}

Write-Output "Running PasteRelinkCrossReference source structure checks"
$moduleText = Read-VbaSource -Path $modulePath
$macroRegistryText = Read-VbaSource -Path $macroRegistryPath
$ribbonText = Get-Content -LiteralPath $ribbonPath -Raw

Assert-True (Test-Path -LiteralPath $modulePath) "Module should exist: $modulePath"
Assert-True ($moduleText.Contains("Public Sub Core_PasteRelinkCrossReference()")) `
    "Module should expose Core_PasteRelinkCrossReference."
Assert-True ($moduleText.Contains("Public Sub Wrapper_PasteRelinkCrossReference()")) `
    "Module should expose the parameterless wrapper for Application.Run."
Assert-True ($moduleText.Contains("sel.PasteAndFormat wdFormatOriginalFormatting")) `
    "Module should paste with original formatting to keep fields."
Assert-True ($moduleText.Contains("Set pastedFields = CollectPastedFields(pastedRange)")) `
    "Module should capture pasted field objects before rewriting field codes."
Assert-True ($moduleText.Contains("UpdateCollectedFields pastedFields")) `
    "Module should update exactly the fields collected from the pasted range."
Assert-True ($moduleText.Contains("StartCustomRecord")) `
    "Module should wrap paste and relink into one undo record."
Assert-True ($macroRegistryText.Contains("Wrapper_PasteRelinkCrossReference")) `
    "Macro registry should register Wrapper_PasteRelinkCrossReference."
Assert-True ($ribbonText.Contains('onAction="PasteRelinkCrossReference"')) `
    "Ribbon should wire a button to PasteRelinkCrossReference."

Write-Output "Source checks PASS"

function Get-TableLabel($doc, $field) {
    # 从 REF 域代码里取书签名
    $code = $field.Code.Text -replace "`r", ' '
    if ($code.Trim() -match '^(?i)(REF|PAGEREF)\s+(\S+)') { return $Matches[2] }
    return ''
}

$word = $null
$doc = $null
try {
    $word = Start-RAToolsWordSession
    $doc = $word.Documents.Add()
    [void]$doc.VBProject.VBComponents.Import($modulePath)

    # 确保有 Table 题注标签
    $hasTable = $false
    foreach ($label in $word.CaptionLabels) { if ($label.Name -eq 'Table') { $hasTable = $true } }
    if (-not $hasTable) { [void]$word.CaptionLabels.Add('Table') }

    # 静默模式：出错时不弹 MsgBox，避免无人值守测试卡在对话框上
    [void]$word.Run('EnablePasteRelinkQuietMode')

    $sel = $word.Selection

    # 第 1 段：正文；第 2 段：题注
    $sel.TypeText('This is body text, see ')
    $sel.TypeParagraph()
    $sel.InsertCaption('Table', ' Example Header') | Out-Null
    $sel.TypeParagraph()

    # 在第 1 段末尾插入指向题注的交叉引用，再补句号
    $p1 = $doc.Paragraphs.Item(1).Range
    $p1.End = $p1.End - 1
    $p1.Collapse(0)
    $p1.InsertCrossReference('Table', 3, 1, $true) | Out-Null

    $refField = $doc.Paragraphs.Item(1).Range.Fields.Item(1)
    $sel.SetRange($refField.Result.End, $refField.Result.End)
    $sel.TypeText('.')

    $doc.Fields.Update() | Out-Null
    $sourceText = $doc.Content.Text
    Assert-True ($sourceText -match 'see\s+Table 1\.') "Source body should reference Table 1. Actual: $sourceText"

    # 复制两段，移动到文末
    $doc.Range($doc.Paragraphs.Item(1).Range.Start, $doc.Paragraphs.Item(2).Range.End).Select()
    $sel.Copy()
    Start-Sleep -Milliseconds 300

    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
    [void]$word.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $word, @("Core_PasteRelinkCrossReference"))
    Start-Sleep -Milliseconds 300

    $resultText = $doc.Content.Text
    Write-Output "Result text: $resultText"

    # 粘贴出来的正文引用应指向 Table 2，题注也应是 Table 2
    $pastedPart = $resultText.Substring($resultText.LastIndexOf("This is body text"))
    Assert-True ($pastedPart -match 'This is body text, see\s+Table 2\.') `
        "Pasted body text should reference Table 2. Actual: $pastedPart"
    Assert-True ($pastedPart -match 'Table 2 Example Header') `
        "Pasted caption should read Table 2 Example Header. Actual: $pastedPart"

    # 原有正文引用必须保持不变
    Assert-True ($resultText -match '^This is body text, see\s+Table 1\.') `
        "Original body text should still reference Table 1. Actual: $resultText"

    # 全文只有两条 REF：原文一条、粘贴块一条，且粘贴块的那条指向新书签
    $refNames = @()
    $refResults = @()
    for ($i = 1; $i -le $doc.Fields.Count; $i++) {
        $f = $doc.Fields.Item($i)
        if ($f.Type -eq 3) {
            $refNames += (Get-TableLabel $doc $f)
            $refResults += ($f.Result.Text -replace "`r", ' ').Trim()
        }
    }
    Assert-True ($refNames.Count -eq 2) "Document should contain exactly two REF fields, found $($refNames.Count)."
    Assert-True ($refResults[1] -eq 'Table 2') "Pasted REF should resolve to Table 2, got '$($refResults[1])'."
    Assert-True ($refResults[0] -eq 'Table 1') "Original REF should still resolve to Table 1, got '$($refResults[0])'."
    Assert-True ($refNames[0] -ne $refNames[1]) `
        "Pasted REF should point at a new bookmark, both point at '$($refNames[0])'."

    # 撤销一次应回到粘贴前的状态
    [void]$doc.Undo()
    $undoneText = $doc.Content.Text
    Assert-True (-not ($undoneText -match 'see\s+Table 2\.')) `
        "Single undo should remove the whole paste+relink operation. Actual: $undoneText"

    Write-Output "COM behavior checks PASS"
    Write-Output "PASS PasteRelinkCrossReference.Tests"
}
finally {
    if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
    Stop-RAToolsWordSession -Word $word
}
