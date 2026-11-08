$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
$word=$null; $doc=$null
function Assert-True($condition,$message) { if(-not $condition){throw $message} }
try {
    $word=Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $doc=$word.Documents.Add()
    [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Engine_StyleNames.bas'))
    $component=$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Macro_RemoveUnusedStyles.bas'))
    $code=$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
    $code=[regex]::Replace($code,'\bMsgBox\b','CaptureStyleMessage')
    $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
    $component.CodeModule.AddFromString($code)
    $component.CodeModule.AddFromString(@'
Private Function CaptureStyleMessage(ByVal message As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
End Function
Public Function UnusedStyleGoneProbe() As Boolean
    UnusedStyleGoneProbe = Not StyleExists(ActiveDocument, "Unused-F")
End Function
'@)
    $used=@('Body-F','HeaderOnly-F','FooterOnly-F','FootnoteOnly-F','TextboxOnly-F','SecondHeader-F')
    foreach($name in ($used+@('Unused-F'))) { [void]$doc.Styles.Add($name,1) }
    $doc.Content.Text='Body text'
    $doc.Content.Style='Body-F'
    $header=$doc.Sections.Item(1).Headers.Item(1).Range
    $header.Text='Header only'; $header.Style='HeaderOnly-F'
    $footer=$doc.Sections.Item(1).Footers.Item(1).Range
    $footer.Text='Footer only'; $footer.Style='FooterOnly-F'
    $footnote=$doc.Footnotes.Add($doc.Range(0,0),[Type]::Missing,'Footnote only')
    $footnote.Range.Style='FootnoteOnly-F'
    $shape=$doc.Shapes.AddTextbox(1,40,40,160,80)
    $shape.TextFrame.TextRange.Text='Textbox only'
    $shape.TextFrame.TextRange.Style='TextboxOnly-F'
    $position=$doc.Content.End-1
    $doc.Range($position,$position).InsertBreak(2)
    $secondHeader=$doc.Sections.Item(2).Headers.Item(1)
    $secondHeader.LinkToPrevious=$false
    $secondHeader.Range.Text='Second header only'
    $secondHeader.Range.Style='SecondHeader-F'
    [void]$doc.Styles.Add('HeaderCharacter-F',2)
    $character=$secondHeader.Range.Characters.Item(1)
    $character.Style='HeaderCharacter-F'
    [void]$doc.Styles.Add('FooterTable-F',3)
    $tableRange=$doc.Sections.Item(1).Footers.Item(1).Range.Duplicate
    $tableRange.Collapse(0)
    $table=$doc.Tables.Add($tableRange,1,1)
    $table.Style='FooterTable-F'
    $used += @('HeaderCharacter-F','FooterTable-F')
    $base=$doc.Styles.Add('Base-F',1)
    $child=$doc.Styles.Add('Child-F',1)
    $child.BaseStyle='Base-F'
    $doc.Paragraphs.Item(1).Range.Style='Child-F'
    $doc.Paragraphs.Item(2).Range.Style='Body-F'
    $headerShape=$doc.Sections.Item(1).Headers.Item(1).Shapes.AddTextbox(1,20,20,140,60)
    [void]$doc.Styles.Add('HeaderTextbox-F',1)
    $headerShape.TextFrame.TextRange.Text='Header textbox'
    $headerShape.TextFrame.TextRange.Style='HeaderTextbox-F'
    $used += @('Base-F','Child-F','HeaderTextbox-F')
    $word.Run('RemoveUnusedStyles')
    foreach($name in $used) {
        $found=$null
        try{$found=$doc.Styles.Item($name)}catch{}
        Assert-True ($null -ne $found) "Used style must survive across stories: $name"
    }
    Assert-True ($word.Run('UnusedStyleGoneProbe')) 'Actually unused style should be removed'
    Assert-True ($header.Style.NameLocal -eq 'HeaderOnly-F') 'Header formatting reference must remain'
    Assert-True ($footnote.Range.Style.NameLocal -eq 'FootnoteOnly-F') 'Footnote formatting reference must remain'
    Assert-True ($shape.TextFrame.TextRange.Style.NameLocal -eq 'TextboxOnly-F') 'Textbox formatting reference must remain'
    Write-Host 'PASS styles in body, headers, footer, footnote and textbox survive; unused style removed'
} finally {
    if($null -ne $doc){$doc.Close(0)}
    Stop-RAToolsWordSession -Word $word
}
Write-Host 'PASS UnusedStylesSafety.Tests'
