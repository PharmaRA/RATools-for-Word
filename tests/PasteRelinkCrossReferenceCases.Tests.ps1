#requires -version 5.1
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

# Macro_PasteRelinkCrossReference 多场景验证：
#   1. 只粘贴题注本身（粘贴块内没有任何引用）-> 不做重连
#   2. 整条题注引用与仅编号引用，两条引用共用同一个新书签
#   3. 粘贴块不含题注 -> 引用保持原样（退化为普通粘贴）
#   4. 粘贴到文档中间 -> 只重连粘贴块内的引用

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $repoRoot "modules\Macro_PasteRelinkCrossReference.bas"

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Invoke-RelinkPaste {
    param($Word)
    [void]$Word.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $Word, @("Core_PasteRelinkCrossReference"))
    Start-Sleep -Milliseconds 250
}

function Get-RefFields {
    param($Doc)
    $items = @()
    for ($i = 1; $i -le $Doc.Fields.Count; $i++) {
        $f = $Doc.Fields.Item($i)
        if ($f.Type -eq 3 -or $f.Type -eq 37) {
            $code = ($f.Code.Text -replace "`r", ' ').Trim()
            $name = ''
            if ($code -match '^(?i)(REF|PAGEREF)\s+"?([^"\s]+)"?') { $name = $Matches[2] }
            $items += [pscustomobject]@{
                Name = $name
                Code = $code
                Result = ($f.Result.Text -replace "`r", ' ').Trim()
            }
        }
    }
    return $items
}

function New-TestDocument {
    param($Word)
    $doc = $Word.Documents.Add()
    [void]$doc.VBProject.VBComponents.Import($modulePath)
    foreach ($name in @('Table', 'Figure')) {
        $exists = $false
        foreach ($label in $Word.CaptionLabels) { if ($label.Name -eq $name) { $exists = $true } }
        if (-not $exists) { [void]$Word.CaptionLabels.Add($name) }
    }
    # 静默模式：出错时不弹 MsgBox，避免无人值守测试卡在对话框上
    [void]$Word.Run('EnablePasteRelinkQuietMode')
    return $doc
}

$word = $null
$docs = @()
try {
    $word = Start-RAToolsWordSession

    # ---------- 场景 1：粘贴块里只有题注，没有任何引用 ----------
    Write-Output "Case 1: paste caption only (no references in block)"
    $doc = New-TestDocument $word
    $docs += $doc
    $sel = $word.Selection
    $sel.InsertCaption('Table', ' Alpha Header') | Out-Null
    $sel.TypeParagraph()
    $sel.InsertCaption('Table', ' Beta Header') | Out-Null
    $doc.Fields.Update() | Out-Null
    $refsBefore = (Get-RefFields $doc).Count

    $doc.Paragraphs.Item(2).Range.Select()
    $sel.Copy()
    Start-Sleep -Milliseconds 250
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
    Invoke-RelinkPaste $word

    $text = $doc.Content.Text
    Write-Output "  text: $text"
    Assert-True ($text -match 'Table 3 Beta Header') "Case 1: pasted caption should be renumbered to Table 3. Actual: $text"
    Assert-True ((Get-RefFields $doc).Count -eq $refsBefore) "Case 1: no references should be added or removed."

    # ---------- 场景 2：整条题注引用 + 仅编号引用 ----------
    Write-Output "Case 2: entire caption and number-only references"
    $doc = New-TestDocument $word
    $docs += $doc
    $sel = $word.Selection
    $sel.InsertCaption('Table', ' Gamma Header') | Out-Null
    $sel.TypeParagraph()
    $sel.TypeText('Whole: ')
    $p = $doc.Paragraphs.Item(2).Range
    $p.End = $p.End - 1
    $p.Collapse(0)
    $p.InsertCrossReference('Table', 2, 1, $true) | Out-Null
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
    $sel.TypeText(' / Number: ')
    $p = $doc.Paragraphs.Item(2).Range
    $p.End = $p.End - 1
    $p.Collapse(0)
    $p.InsertCrossReference('Table', 3, 1, $true) | Out-Null
    $doc.Fields.Update() | Out-Null

    $doc.Range($doc.Paragraphs.Item(1).Range.Start, $doc.Paragraphs.Item(2).Range.End).Select()
    $sel.Copy()
    Start-Sleep -Milliseconds 250
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
    Invoke-RelinkPaste $word

    $text = $doc.Content.Text
    Write-Output "  text: $text"
    # 粘贴块里的两条引用都应落在新题注上：整条题注引用含 "Table 2"，仅编号引用也是 "Table 2"
    $pastedText = $text.Substring($text.IndexOf('/ Number: Table 1') + 1)
    Assert-True ($pastedText -match 'Table 2 Gamma Header') `
        "Case 2: pasted whole-caption reference should resolve to Table 2. Actual: $pastedText"
    Assert-True ($pastedText -match '/ Number: Table 2') `
        "Case 2: pasted number-only reference should resolve to Table 2. Actual: $pastedText"
    # Word 会为「整条题注」和「仅编号」各建一个书签，二者本就不同；
    # 断言的是：粘贴后的两条引用都改指新书签，且与原文的两个书签都不同。
    $refs = Get-RefFields $doc
    Assert-True ($refs.Count -eq 4) "Case 2: expected 4 REF fields, found $($refs.Count)."
    $originalNames = @($refs[0].Name, $refs[1].Name)
    $pastedNames = @($refs[2].Name, $refs[3].Name)
    Assert-True ($pastedNames[0] -notin $originalNames) `
        "Case 2: pasted whole-caption ref must point at a new bookmark, got '$($pastedNames[0])'."
    Assert-True ($pastedNames[1] -notin $originalNames) `
        "Case 2: pasted number-only ref must point at a new bookmark, got '$($pastedNames[1])'."
    Assert-True ($refs[0].Result -eq 'Table 1 Gamma Header') "Case 2: original whole-caption ref unchanged, got '$($refs[0].Result)'."
    Assert-True ($refs[1].Result -eq 'Table 1') "Case 2: original number ref unchanged, got '$($refs[1].Result)'."

    # ---------- 场景 3：粘贴块里没有题注 ----------
    Write-Output "Case 3: block without caption"
    $doc = New-TestDocument $word
    $docs += $doc
    $sel = $word.Selection
    $sel.InsertCaption('Table', ' Delta Header') | Out-Null
    $sel.TypeParagraph()
    $sel.TypeText('See ')
    $p = $doc.Paragraphs.Item(2).Range
    $p.End = $p.End - 1
    $p.Collapse(0)
    $p.InsertCrossReference('Table', 3, 1, $true) | Out-Null
    $doc.Fields.Update() | Out-Null

    $doc.Paragraphs.Item(2).Range.Select()
    $sel.Copy()
    Start-Sleep -Milliseconds 250
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
    Invoke-RelinkPaste $word

    $refs = Get-RefFields $doc
    Assert-True ($refs.Count -eq 2) "Case 3: expected 2 REF fields, found $($refs.Count)."
    Assert-True ($refs[0].Name -eq $refs[1].Name) `
        "Case 3: without a pasted caption the pasted ref keeps the original bookmark."
    Assert-True ($refs[1].Result -eq 'Table 1') "Case 3: pasted ref should still resolve to Table 1, got '$($refs[1].Result)'."

    # ---------- 场景 4：粘贴到文档中间 ----------
    Write-Output "Case 4: paste in the middle of the document"
    $doc = New-TestDocument $word
    $docs += $doc
    $sel = $word.Selection
    $sel.InsertCaption('Table', ' Epsilon Header') | Out-Null
    $sel.TypeParagraph()
    $sel.TypeText('See ')
    $p = $doc.Paragraphs.Item(2).Range
    $p.End = $p.End - 1
    $p.Collapse(0)
    $p.InsertCrossReference('Table', 3, 1, $true) | Out-Null
    $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
    $sel.TypeParagraph()
    $sel.TypeText('TAIL LINE')
    $doc.Fields.Update() | Out-Null

    $doc.Range($doc.Paragraphs.Item(1).Range.Start, $doc.Paragraphs.Item(2).Range.End).Select()
    $sel.Copy()
    Start-Sleep -Milliseconds 250
    $target = $doc.Paragraphs.Item(3).Range.Start
    $sel.SetRange($target, $target)
    Invoke-RelinkPaste $word

    $text = $doc.Content.Text
    Write-Output "  text: $text"
    Assert-True ($text -match 'TAIL LINE') "Case 4: tail paragraph should survive."
    Assert-True ($text -match 'Table 2 Epsilon Header') "Case 4: pasted caption should be Table 2. Actual: $text"
    Assert-True ($text -match 'See Table 2') "Case 4: pasted ref should point at Table 2. Actual: $text"
    $refs = Get-RefFields $doc
    Assert-True ($refs.Count -eq 2) "Case 4: expected 2 REF fields, found $($refs.Count)."
    Assert-True ($refs[0].Result -eq 'Table 1') "Case 4: original ref should stay Table 1, got '$($refs[0].Result)'."
    Assert-True ($refs[1].Result -eq 'Table 2') "Case 4: pasted ref should be Table 2, got '$($refs[1].Result)'."

    Write-Output "PASS PasteRelinkCrossReferenceCases.Tests"
}
finally {
    foreach ($doc in $docs) { try { $doc.Close(0) } catch { } }
    Stop-RAToolsWordSession -Word $word
}
