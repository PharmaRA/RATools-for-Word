param([switch]$UseWord = $true)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$sourcePath = Join-Path $repo 'modules\Macro_AcceptFormatRevisions.bas'
$source = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::GetEncoding(936))
$testRoot = Join-Path (Join-Path $repo 'backups') ('RATools_FormatRepair_' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName Microsoft.VisualBasic

function Assert-True($condition, $message) {
    if (-not $condition) { throw $message }
}
function New-Package($path, [bool]$badXml) {
    $zip = [IO.Compression.ZipFile]::Open($path, 'Create')
    try {
        $document = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
            '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange w:id="10" w:original="%1."/></w:numPr><w:rPr><w:ins w:id="99"/></w:rPr><w:pPrChange w:id="1"/></w:pPr>' +
            '<w:r><w:rPr><w:rPrChange w:id="2"/></w:rPr><w:t>Keep text</w:t></w:r>' +
            '<w:ins w:id="3"><w:r><w:t>Keep insertion</w:t></w:r></w:ins></w:p>' +
            '<w:tbl><w:tblPr><w:tblPrChange w:id="11"><w:tblPr/></w:tblPrChange></w:tblPr>' +
            '<w:tblGrid><w:gridCol w:w="2000"/><w:tblGridChange w:id="12"><w:tblGrid/></w:tblGridChange></w:tblGrid>' +
            '<w:tr><w:tblPrEx><w:tblPrExChange w:id="13"><w:tblPrEx/></w:tblPrExChange></w:tblPrEx>' +
            '<w:trPr><w:ins w:id="98"/><w:trPrChange w:id="14"><w:trPr/></w:trPrChange></w:trPr>' +
            '<w:tc><w:tcPr><w:tcPrChange w:id="15"><w:tcPr/></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
            '<w:ins w:id="16"><w:r><w:instrText> DATE \@ "yyyy-MM-dd" </w:instrText></w:r></w:ins>' +
            '<w:del w:id="17"><w:r><w:delInstrText> DATE </w:delInstrText></w:r></w:del>' +
            '<w:ins w:id="18"><w:r><w:rPr/></w:r></w:ins>' +
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
            '<w:ins w:id="19"><w:r><w:t>Keep result</w:t></w:r></w:ins>' +
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
            '<w:ins w:id="20"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins></w:p>' +
            '<w:sectPr><w:sectPrChange w:id="21"><w:sectPr/></w:sectPrChange></w:sectPr>' +
            '</w:body></w:document>'
        if ($badXml) { $document = '<w:document broken' }
        $parts = @{
            '[Content_Types].xml' = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
            '_rels/.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
            'word/document.xml' = $document
            'word/header1.xml' = '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pPrChange w:id="4"/></w:pPr></w:p></w:hdr>'
        }
        foreach ($name in $parts.Keys) {
            $writer = [IO.StreamWriter]::new($zip.CreateEntry($name).Open())
            try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
        }
        $stream = $zip.CreateEntry('word/media/probe.bin').Open()
        try { $stream.Write([byte[]](0..255), 0, 256) } finally { $stream.Dispose() }
    } finally { $zip.Dispose() }
}
function Invoke-Processor($name, [bool]$badXml, [bool]$lockTarget = $false, [bool]$collision = $false) {
    $work = Join-Path $testRoot $name
    [void][IO.Directory]::CreateDirectory($work)
    $path = Join-Path $work 'test.docx'
    New-Package $path $badXml
    $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $payload = Join-Path $work 'work'
    [void][IO.Directory]::CreateDirectory($payload)
    $resultPath = Join-Path $work 'result.txt'
    $ready = $path + '.ratools-work.tmp'
    if ($collision) { [IO.File]::WriteAllText($ready, 'existing transaction material') }
    $lock = $null
    try {
        if ($lockTarget) { $lock = [IO.File]::Open($path, 'Open', 'Read', 'Read') }
        if ($script:processorExtension -eq '.vbs') {
            $output = & cscript.exe //nologo $script:processorPath $path $payload 2>&1
        } else {
            $output = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $script:processorPath -DocumentPath $path -WorkDirectory $payload -ResultPath $resultPath 2>&1
        }
        $code = $LASTEXITCODE
    } finally { if ($null -ne $lock) { $lock.Dispose() } }
    if ($badXml -or $lockTarget -or $collision) {
        Assert-True ($code -ne 0) "$name must fail instead of accepting an invalid/unwritable package"
        Assert-True ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -eq $before) "$name must preserve original bytes"
        if ($collision) {
            Assert-True (Test-Path -LiteralPath $ready) 'Existing transaction file must stay at its original path'
            Assert-True ([IO.File]::ReadAllText($ready) -eq 'existing transaction material') 'Existing transaction bytes must survive'
        }
    } else {
        Assert-True ($code -eq 0) "$name processor failed: $output"
        Assert-True ((Get-FileHash -LiteralPath ($ready + '.bak') -Algorithm SHA256).Hash -eq $before) 'Transaction backup must equal the original bytes'
        $zip = [IO.Compression.ZipFile]::OpenRead($path)
        try {
            Assert-True ($zip.Entries.Count -eq 5) 'All package parts must survive'
            $reader = [IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            Assert-True ($xml.SelectNodes("//*[local-name()='rPrChange' or local-name()='pPrChange' or local-name()='tblPrChange' or local-name()='tblGridChange' or local-name()='tblPrExChange' or local-name()='trPrChange' or local-name()='tcPrChange' or local-name()='sectPrChange' or local-name()='numberingChange']").Count -eq 0) 'All format changes must be accepted'
            $remainingInsIds = @($xml.SelectNodes("//*[local-name()='ins']") | ForEach-Object { $_.GetAttribute('id', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') })
            Assert-True (($remainingInsIds -join ',') -eq '99,3,98,19,20') "Content and whole-field insertions must survive; got: $($remainingInsIds -join ',')"
            Assert-True ($xml.SelectNodes("//*[local-name()='del']").Count -eq 0) 'Field code deletion must be accepted'
            Assert-True ($xml.SelectNodes("//*[local-name()='delInstrText']").Count -eq 0) 'Deleted field code instruction must be removed'
            $instrText = (@($xml.SelectNodes("//*[local-name()='instrText']") | ForEach-Object { $_.InnerText }) -join '|')
            Assert-True ($instrText -eq ' DATE \@ "yyyy-MM-dd" | PAGE ') "Accepted and preserved field codes must survive: $instrText"
            $wordText = (@($xml.SelectNodes("//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join '')
            Assert-True ($wordText -eq 'Keep textKeep insertionCellKeep result1') "Word text must survive: $wordText"
            $reader = [IO.StreamReader]::new($zip.GetEntry('word/header1.xml').Open())
            try { [xml]$headerXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            Assert-True ($headerXml.SelectNodes("//*[local-name()='pPrChange']").Count -eq 0) 'Header formatting revision must be accepted'
            $stream = $zip.GetEntry('word/media/probe.bin').Open()
            $memory = [IO.MemoryStream]::new()
            try { $stream.CopyTo($memory); $bytes = $memory.ToArray() } finally { $stream.Dispose(); $memory.Dispose() }
            Assert-True ([Convert]::ToBase64String($bytes) -eq [Convert]::ToBase64String([byte[]](0..255))) 'Binary payload must survive byte for byte'
        } finally { $zip.Dispose() }
    }
    Write-Host "PASS $name"
}

try {
    # Exercise the actual emitted processor, without depending on a copied algorithm.
    $body = [regex]::Match($source, '(?s)Private Function WriteStripScript\(.*?End Function').Value
    $emitted = @([regex]::Matches($body, '(?m)^\s*Print #f, "((?:""|[^"])*)"\s*$') | ForEach-Object { $_.Groups[1].Value.Replace('""','"') })
    Assert-True ($emitted.Count -gt 10) 'Processor could not be extracted'
    $script:processorExtension = '.vbs'
    if (($emitted -join "`n") -match '(?m)^param\(') { $script:processorExtension = '.ps1' }
    $script:processorPath = Join-Path $testRoot ('processor' + $script:processorExtension)
    $encoding = [Text.UTF8Encoding]::new($true)
    if ($script:processorExtension -eq '.vbs') { $encoding = [Text.Encoding]::ASCII }
    [IO.File]::WriteAllLines($script:processorPath, $emitted, $encoding)
    Invoke-Processor 'MalformedXmlPreservesOriginal' $true
    Invoke-Processor 'CompletePackageAndContentRevisionsSurvive' $false
    Invoke-Processor 'FailedReplacementPreservesOriginal' $false $true
    Invoke-Processor 'ExistingTransactionPreserved' $false $false $true
    if ($UseWord) {
        Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
        $word = $null; $doc = $null
        try {
            $word = Start-RAToolsWordSession
            # This dedicated instance must not run startup templates' UI events.
            $word.AddIns.Unload($false)
            $doc = $word.Documents.Add()
            [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Core_Session.bas'))
            [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Core_Files.bas'))
            $component = $doc.VBProject.VBComponents.Import($sourcePath)
            $component.CodeModule.AddFromString("Public Function RepairCompileProbe() As Boolean`r`nRepairCompileProbe = True`r`nEnd Function")
            Assert-True ($word.Run('RepairCompileProbe') -eq $true) 'Changed VBA module must compile in Word'
            Write-Host 'PASS Word module compile'
            $realPath = Join-Path $testRoot 'Word roundtrip.docx'
            $realWork = Join-Path $testRoot 'WordWork'
            [void][IO.Directory]::CreateDirectory($realWork)
            $realScript = Join-Path $realWork 'strip.ps1'
            $realResult = Join-Path $realWork 'result.txt'
            Write-Host 'Creating Word revision fixture'
            $realDoc = $word.Documents.Add()
            try {
                $realDoc.Content.Text = "Original text`r`n"
                $tblRange = $realDoc.Content; $tblRange.Collapse(0)
                $tbl = $realDoc.Tables.Add($tblRange, 2, 2)
                $tbl.Cell(1, 1).Range.Text = 'A1'
                $tbl.Cell(1, 2).Range.Text = 'B1'
                $tbl.Cell(2, 1).Range.Text = 'A2'
                $tbl.Cell(2, 2).Range.Text = 'B2'
                $fldRange = $realDoc.Content; $fldRange.Collapse(0)
                $fldRange.InsertAfter("`r`nRef: ")
                $fldRange.Collapse(0)
                $realDoc.Bookmarks.Add('bm1', $realDoc.Paragraphs[1].Range) | Out-Null
                $fld = $realDoc.Fields.Add($fldRange, 3, 'bm1 \h', $true)
                $hlRange = $realDoc.Content; $hlRange.Collapse(0)
                $hlRange.InsertAfter("`r`nLink: ")
                $hlRange.Collapse(0)
                $hl = $realDoc.Hyperlinks.Add($hlRange, 'https://a.example.com', $null, $null, 'Click')

                $realDoc.TrackRevisions = $true
                # Formatting, table, section, and field code revisions (must all be accepted)
                $realDoc.Range(0, 8).Font.Bold = 1
                $realDoc.Paragraphs[1].Format.SpaceAfter = 12
                $tbl.Borders.Enable = 1
                $tbl.Cell(1, 1).Shading.BackgroundPatternColor = 13434879
                $tbl.Rows[1].Height = 30
                $tbl.Columns[1].Width = 120
                $realDoc.PageSetup.TopMargin = 85
                $fld.Code.Text = $fld.Code.Text + ' \* MERGEFORMAT '
                $hl.Address = 'https://b.example.com'
                # Content revisions (must all survive)
                $realDoc.Range(13, 13).InsertAfter(' inserted text')
                $tbl.Cell(2, 2).Range.Text = 'B2 edited'
                $newFldRange = $realDoc.Content; $newFldRange.Collapse(0)
                $newFldRange.InsertAfter("`r`nNew: ")
                $newFldRange.Collapse(0)
                [void]$realDoc.Fields.Add($newFldRange, -1, 'PAGE', $true)

                $expectedText = $realDoc.Content.Text
                Assert-True ($realDoc.Revisions.Count -ge 8) 'Fixture must contain format, table, section, field code, and content revisions'
                Write-Host 'Saving Word revision fixture'
                $realDoc.SaveAs2([string]$realPath, 12)
            } finally { $realDoc.Close(0); $realDoc = $null }
            $component.CodeModule.AddFromString(@"
Public Function RepairRunProbe() As String
    Dim removed As Long, errorText As String
    If Not WriteStripScript("$realScript") Then
        RepairRunProbe = "EMIT_FAILED"
    ElseIf RunStripScript("$realScript", "$realPath", "$realWork", "$realResult", removed, errorText) Then
        RepairRunProbe = "OK=" & removed
    Else
        RepairRunProbe = "UNAVAILABLE=" & errorText
    End If
End Function
"@)
            Write-Host 'Running VBA processor bridge'
            $originalHash = (Get-FileHash -LiteralPath $realPath -Algorithm SHA256).Hash
            $bridgeResult = $word.Run('RepairRunProbe')
            if ($bridgeResult -notlike 'OK=*') {
                Assert-True ($bridgeResult -like 'UNAVAILABLE=*') 'VBA must emit the actual processor'
                Assert-True ((Get-FileHash -LiteralPath $realPath -Algorithm SHA256).Hash -eq $originalHash) 'Unavailable launch must not modify the document'
                Write-Host "LIMITATION Word could not launch processor: $bridgeResult"
                & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $realScript -DocumentPath $realPath -WorkDirectory $realWork -ResultPath $realResult
                Assert-True ($LASTEXITCODE -eq 0) 'Externally launched VBA-emitted processor must succeed'
            }
            $realDoc = $word.Documents.Open($realPath)
            try {
                Assert-True ($realDoc.Content.Text -eq $expectedText) 'Word roundtrip text must survive'
                $formatCount = 0; $contentCount = 0
                foreach ($revision in $realDoc.Revisions) {
                    $inFieldCode = $false
                    if ($revision.Type -in @(1, 2) -and $revision.Range.Fields.Count -eq 0) {
                        foreach ($f in $revision.Range.Paragraphs[1].Range.Fields) {
                            if ($revision.Range.Start -ge $f.Code.Start -and $revision.Range.End -le $f.Code.End) {
                                $inFieldCode = $true; break
                            }
                        }
                    }
                    if ($revision.Type -in @(3, 4, 5, 8, 10, 11, 12, 13) -or $inFieldCode) { $formatCount++ } else { $contentCount++ }
                }
                Assert-True ($formatCount -eq 0) "Word must report zero formatting/table/section/field-code revisions; found $formatCount"
                Assert-True ($contentCount -ge 3) "Word must still report the content revisions; found $contentCount"
                Write-Host 'PASS VBA-emitted processor and Word document reopen'
            } finally { $realDoc.Close(0) }
            # Inject only the unavailable process boundary and interactive message box.
            # The production save/close/reopen and Word API fallback execute unchanged.
            $fallbackPath = Join-Path $testRoot 'Fallback.docx'
            [IO.File]::Copy(($realPath + '.ratools-WordWork.tmp.bak'), $fallbackPath)
            $code = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
            $code = [regex]::Replace($code, '\bMsgBox\b', 'CaptureRepairMessage')
            $code = [regex]::Replace($code, '(?s)(Private Function RunStripScript\(.*?\) As Boolean).*?End Function', '$1' + "`r`n    errText = ""Injected launch failure""`r`nEnd Function")
            $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
            $component.CodeModule.AddFromString($code)
            $component.CodeModule.AddFromString(@"
Private Function CaptureRepairMessage(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    CaptureRepairMessage = vbNo
End Function
Public Sub RepairFallbackProbe()
    Documents("Fallback.docx").Activate
    AcceptFormatRevisionsByOoxml "$fallbackPath"
End Sub
"@)
            $fallbackDoc = $word.Documents.Open($fallbackPath)
            $doc.Activate()
            $word.Run('RepairFallbackProbe')
            $fallbackDoc = $word.ActiveDocument
            try {
                Assert-True ($fallbackDoc.FullName -eq $fallbackPath) 'Fallback must keep the intended document'
                Assert-True ($fallbackDoc.Content.Text -eq $expectedText) 'Fallback must preserve text'
                $types = @($fallbackDoc.Revisions | ForEach-Object { [int]$_.Type })
                $fallbackFormatLeft = 0
                foreach ($revision in $fallbackDoc.Revisions) {
                    $inFieldCode = $false
                    if ($revision.Type -in @(1, 2) -and $revision.Range.Fields.Count -eq 0) {
                        foreach ($f in $revision.Range.Paragraphs[1].Range.Fields) {
                            if ($revision.Range.Start -ge $f.Code.Start -and $revision.Range.End -le $f.Code.End) {
                                $inFieldCode = $true; break
                            }
                        }
                    }
                    if ($revision.Type -in @(3, 4, 5, 8, 10, 11, 12, 13) -or $inFieldCode) { $fallbackFormatLeft++ }
                }
                Assert-True ($fallbackFormatLeft -eq 0) "Fallback must accept all format/table/section/field-code revisions; left $fallbackFormatLeft"
                Assert-True ($types.Count -ge 3) "Fallback must retain content revisions; left $($types.Count)"
                Write-Host 'PASS Word fallback after unavailable processor'
                $cancelClass = $doc.VBProject.VBComponents.Add(2)
                $cancelClass.Name = 'RepairCancelEvents'
                $cancelClass.CodeModule.AddFromString(@'
Public WithEvents App As Word.Application
Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    Cancel = True
End Sub
'@)
                $component.CodeModule.AddFromString(@"
Public Sub RepairCancelSaveProbe()
    Dim cancelEvents As RepairCancelEvents
    Set cancelEvents = New RepairCancelEvents
    Set cancelEvents.App = Application
    Documents("Fallback.docx").Activate
    AcceptFormatRevisionsByOoxml "$fallbackPath"
    Set cancelEvents.App = Nothing
End Sub
"@)
                $fallbackDoc.Content.InsertAfter(' pending unsaved change')
                $pendingText = $fallbackDoc.Content.Text
                $openCount = $word.Documents.Count
                $doc.Activate()
                $word.Run('RepairCancelSaveProbe')
                Assert-True ($word.Documents.Count -eq $openCount) 'Cancelled save must leave the document open'
                Assert-True (-not $fallbackDoc.Saved) 'Cancelled save must keep unsaved state'
                Assert-True ($fallbackDoc.Content.Text -eq $pendingText) 'Cancelled save must retain pending edits'
                Write-Host 'PASS Word cancelled save retains open unsaved document'
            } finally { $fallbackDoc.Close(0) }
        } finally {
            if ($null -ne $doc) { $doc.Close(0) }
            Stop-RAToolsWordSession -Word $word
        }
    }
} finally {
    # Recycle only the unique directory created by this test.
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($testRoot, 'OnlyErrorDialogs', 'SendToRecycleBin')
}
Write-Host 'PASS AcceptFormatRevisions.Tests'
