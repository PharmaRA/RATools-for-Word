$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
function Assert-True($condition, $message) { if (-not $condition) { throw $message } }
$word = $null; $doc = $null
try {
    $word = Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $doc = $word.Documents.Add()
    [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Macro_PasteRelinkCrossReference.bas'))
    $word.Run('EnablePasteRelinkQuietMode')
    if (@($word.CaptionLabels | Where-Object { $_.Name -eq 'Table' }).Count -eq 0) { [void]$word.CaptionLabels.Add('Table') }
    $sel = $word.Selection
    $sel.TypeText('References: ')
    $sel.TypeParagraph()
    $sel.InsertCaption('Table', ' Alpha caption') | Out-Null
    $sel.TypeParagraph()
    $sel.InsertCaption('Table', ' Beta caption') | Out-Null
    $sel.TypeParagraph()
    $caption1 = $doc.Paragraphs.Item(2).Range
    $caption2 = $doc.Paragraphs.Item(3).Range
    [void]$doc.Bookmarks.Add('A', $doc.Range($caption1.Start, $caption1.Fields.Item(1).Result.End))
    [void]$doc.Bookmarks.Add('LongCaptionBetaSourceName', $doc.Range($caption2.Start, $caption2.Fields.Item(1).Result.End))
    foreach ($name in @('A','LongCaptionBetaSourceName')) {
        $position = $doc.Paragraphs.Item(1).Range.End - 1
        [void]$doc.Fields.Add($doc.Range($position,$position), -1, ('REF ' + $name + ' \h'))
        $position = $doc.Paragraphs.Item(1).Range.End - 1
        $doc.Range($position,$position).Text = ' / '
    }
    [void]$doc.Fields.Update()
    $sourceText = $doc.Content.Text
    $doc.Range(0, $doc.Paragraphs.Item(3).Range.End).Select()
    $sel.Copy()
    Start-Sleep -Milliseconds 250
    $sel.SetRange($doc.Content.End-1,$doc.Content.End-1)
    $word.Run('Core_PasteRelinkCrossReference')
    $pasteError = $word.Run('GetLastPasteRelinkErrorForTest')
    Assert-True ([string]::IsNullOrEmpty($pasteError)) "Paste must not report an error: $pasteError"
    $references = @($doc.Fields | Where-Object { $_.Type -eq 3 })
    $actual = @($references | ForEach-Object { $_.Result.Text.Trim() })
    Assert-True ($actual.Count -eq 4) 'Expected original and pasted reference pairs'
    Assert-True ($actual[0] -eq 'Table 1' -and $actual[1] -eq 'Table 2') 'Original references must not change'
    Assert-True ($actual[2] -eq 'Table 3') "First pasted reference must point at first pasted caption: $($actual[2])"
    Assert-True ($actual[3] -eq 'Table 4') "Second pasted reference must point at second pasted caption: $($actual[3])"
    [void]$doc.Undo()
    Assert-True ($doc.Content.Text -eq $sourceText) 'One undo must restore the original document text'
    Write-Host 'PASS variable-length reference names preserve all caption targets and undo'
    $doc.Close(0); $doc = $null
    foreach ($description in @(' Same description', '')) {
        $doc = $word.Documents.Add()
        [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Macro_PasteRelinkCrossReference.bas'))
        $word.Run('EnablePasteRelinkQuietMode')
        $sel = $word.Selection
        $sel.TypeText('Reverse references: ')
        $sel.TypeParagraph()
        $sel.InsertCaption('Table', $description) | Out-Null
        $sel.TypeParagraph()
        $sel.InsertCaption('Table', $description) | Out-Null
        $sel.TypeParagraph()
        foreach ($index in @(2,1)) {
            $position = $doc.Paragraphs.Item(1).Range.End - 1
            $doc.Range($position,$position).InsertCrossReference('Table',3,$index,$true) | Out-Null
            $position = $doc.Paragraphs.Item(1).Range.End - 1
            $doc.Range($position,$position).Text = ' / '
        }
        [void]$doc.Fields.Update()
        $sourceReferences = @($doc.Fields | Where-Object { $_.Type -eq 3 } | ForEach-Object { $_.Code.Text.Trim() })
        $doc.Range(0,$doc.Paragraphs.Item(3).Range.End).Select()
        $sel.Copy()
        Start-Sleep -Milliseconds 250
        $sel.SetRange($doc.Content.End-1,$doc.Content.End-1)
        $word.Run('Core_PasteRelinkCrossReference')
        $references = @($doc.Fields | Where-Object { $_.Type -eq 3 })
        Assert-True ($references[2].Code.Text.Trim() -eq $sourceReferences[0]) 'Ambiguous reversed reference must retain its original target'
        Assert-True ($references[3].Code.Text.Trim() -eq $sourceReferences[1]) 'Ambiguous second reference must retain its original target'
        Assert-True (-not [string]::IsNullOrEmpty($word.Run('GetLastPasteRelinkErrorForTest'))) 'Ambiguous matching must report that references were retained'
        Write-Host "PASS ambiguous captions retained with notice (description length $($description.Length))"
        $doc.Close(0); $doc = $null
    }
    $doc = $word.Documents.Add()
    [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Macro_PasteRelinkCrossReference.bas'))
    $word.Run('EnablePasteRelinkQuietMode')
    $sel = $word.Selection
    $sel.InsertCaption('Table', ' Shared description') | Out-Null
    $sel.TypeParagraph()
    $sel.TypeText('References: ')
    $sel.TypeParagraph()
    $sel.InsertCaption('Table', ' Shared description') | Out-Null
    $sel.TypeParagraph()
    foreach ($index in @(1,2)) {
        $position = $doc.Paragraphs.Item(2).Range.End - 1
        $doc.Range($position,$position).InsertCrossReference('Table',3,$index,$true) | Out-Null
        $position = $doc.Paragraphs.Item(2).Range.End - 1
        $doc.Range($position,$position).Text = ' / '
    }
    [void]$doc.Fields.Update()
    $originalRefs = @($doc.Fields | Where-Object { $_.Type -eq 3 } | ForEach-Object { $_.Code.Text.Trim() })
    $doc.Range($doc.Paragraphs.Item(2).Range.Start,$doc.Paragraphs.Item(3).Range.End).Select()
    $sel.Copy()
    Start-Sleep -Milliseconds 250
    $sel.SetRange($doc.Content.End-1,$doc.Content.End-1)
    $word.Run('Core_PasteRelinkCrossReference')
    $refs = @($doc.Fields | Where-Object { $_.Type -eq 3 })
    Assert-True ($refs[2].Code.Text.Trim() -eq $originalRefs[0]) 'External same-description source must not claim the single pasted caption'
    Assert-True ($refs[3].Code.Text.Trim() -eq $originalRefs[1]) 'Ambiguous source identity must preserve the second target'
    Assert-True (-not [string]::IsNullOrEmpty($word.Run('GetLastPasteRelinkErrorForTest'))) 'Source ambiguity must report a notice'
    Write-Host 'PASS multiple source identities cannot claim one pasted caption'
} finally {
    if ($null -ne $doc) { $doc.Close(0) }
    Stop-RAToolsWordSession -Word $word
}
Write-Host 'PASS PasteRelinkSafety.Tests'
