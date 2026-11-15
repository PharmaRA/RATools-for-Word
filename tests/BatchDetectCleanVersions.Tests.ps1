$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\scripts\RATools.Build.psm1') -Force
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$fixtureRoot = Join-Path $repoRoot ('backups\clean-detection-tests-' + [guid]::NewGuid().ToString('N'))
$word = $null
$ownsSession = $false
$script:assertionCount = 0
$script:scenarioCount = 0

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    $script:assertionCount++
    if ($Expected -ne $Actual) { throw "$Message Expected <$Expected>, got <$Actual>." }
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    $script:assertionCount++
    if (-not $Condition) { throw $Message }
}

function Complete-Scenario {
    param([string]$Name)
    $script:scenarioCount++
    Write-Host "PASS $Name"
}

function New-SourceFixture {
    $source = $word.Documents.Add()
    $source.TrackRevisions = $false
    $source.Content.Text = 'Baseline fixture text.'
    return $source
}

function Get-SourceSnapshot {
    param($Document)
    # Establish the probe host before the baseline, outside the detection under test.
    $hostDoc.Activate()
    $storySnapshots = @()
    foreach ($rootStory in $Document.StoryRanges) {
        $story = $rootStory
        do {
            $revisions = @()
            foreach ($revision in $story.Revisions) {
                $revisions += ('{0}:{1}:{2}:{3}' -f $revision.Type, $revision.Range.Start, $revision.Range.End, $revision.Range.Text)
            }
            $storySnapshots += [pscustomobject]@{
                Type = [int]$story.StoryType
                Text = [string]$story.Text
                Revisions = ($revisions -join "`n")
            }
            $story = $story.NextStoryRange
        } while ($null -ne $story)
    }
    $comments = @()
    foreach ($comment in $Document.Comments) {
        $comments += ('{0}:{1}:{2}' -f $comment.Scope.Start, $comment.Scope.End, $comment.Range.Text)
    }
    $fields = @()
    foreach ($field in $Document.Fields) {
        $fields += ('{0}:{1}:{2}:{3}' -f $field.Type, $field.Code.Text, $field.Result.Text, [bool]$field.Locked)
    }
    return [pscustomobject]@{
        Stories = $storySnapshots
        Comments = ($comments -join "`n")
        Fields = ($fields -join "`n")
        Saved = [bool]$Document.Saved
        ReadOnly = [bool]$Document.ReadOnly
        Protection = [int]$Document.ProtectionType
        Tracking = [bool]$Document.TrackRevisions
        FullName = [string]$Document.FullName
    }
}

function Assert-SourceUnchanged {
    param($Before, $Document, [string]$Name)
    $after = Get-SourceSnapshot $Document
    Assert-Equal $Before.Stories.Count $after.Stories.Count "$Name changed the story count."
    for ($index = 0; $index -lt $Before.Stories.Count; $index++) {
        Assert-Equal $Before.Stories[$index].Type $after.Stories[$index].Type "$Name changed a story type."
        Assert-Equal $Before.Stories[$index].Text $after.Stories[$index].Text "$Name changed story text."
        Assert-Equal $Before.Stories[$index].Revisions $after.Stories[$index].Revisions "$Name changed revisions."
    }
    Assert-Equal $Before.Comments $after.Comments "$Name changed comments."
    Assert-Equal $Before.Fields $after.Fields "$Name changed fields or field locks."
    Assert-Equal $Before.Saved $after.Saved "$Name changed the saved state."
    Assert-Equal $Before.ReadOnly $after.ReadOnly "$Name changed read-only state."
    Assert-Equal $Before.Protection $after.Protection "$Name changed document protection."
    Assert-Equal $Before.Tracking $after.Tracking "$Name changed revision tracking."
    Assert-Equal $Before.FullName $after.FullName "$Name changed the document path."
}

function ConvertFrom-ProbeResult {
    param([string]$Value)
    $parts = $Value -split ([char]9), 2
    Assert-Equal 2 $parts.Count 'Probe must return a status and reason.'
    return [pscustomobject]@{ Status = $parts[0]; Reason = $parts[1] }
}

function Invoke-CleanProbe {
    param([string]$Name, [string]$Argument = 'False')
    $hostDoc.Activate()
    $hostDoc.Variables.Item('RAToolsCleanProbeArgument').Value = $Argument
    return $word.Run($Name)
}

function Invoke-DocumentCheck {
    param($Document)
    return ConvertFrom-ProbeResult ([string](Invoke-CleanProbe 'RAToolsCleanDocumentProbe' ([string]$Document.Name)))
}

function Invoke-FileCheck {
    param([string]$Path)
    $hostDoc.Variables.Item('RAToolsCleanFileCompleted').Value = 'False'
    $hostDoc.Variables.Item('RAToolsCleanOpenedReadOnly').Value = 'False'
    $hostDoc.Variables.Item('RAToolsCleanOpenedSecurity').Value = '0'
    $hostDoc.Variables.Item('RAToolsCleanWorkerWindow').Value = '0'
    $result = ConvertFrom-ProbeResult ([string](Invoke-CleanProbe 'RAToolsCleanFileProbe' $Path))
    Assert-Equal 'True' ([string]$hostDoc.Variables.Item('RAToolsCleanFileCompleted').Value) 'File detection did not complete.'
    $workerWindow = [IntPtr]([long]$hostDoc.Variables.Item('RAToolsCleanWorkerWindow').Value)
    if ($workerWindow -ne [IntPtr]::Zero) {
        Assert-Equal $false ([RAToolsCleanTestWindows]::IsWindow($workerWindow)) 'The dedicated inspection Word instance was left open.'
    }
    return $result
}

function Assert-Classification {
    param($Document, [string]$Status, [string[]]$Reasons, [string]$Name, [switch]$KeepOpen)
    $before = Get-SourceSnapshot $Document
    $control = Get-SourceSnapshot $Document
    Assert-Equal $before.Stories.Count $control.Stories.Count "$Name story count is unstable before detection."
    for ($index = 0; $index -lt $before.Stories.Count; $index++) {
        Assert-Equal $before.Stories[$index].Text $control.Stories[$index].Text "$Name story text is unstable before detection."
    }
    $countBefore = $word.Documents.Count
    $result = Invoke-DocumentCheck $Document
    Assert-Equal $Status $result.Status "$Name classification: $($result.Reason)"
    foreach ($reason in $Reasons) {
        Assert-True ($result.Reason.Contains($reason)) "$Name omitted reason <$reason>."
    }
    Assert-Equal $countBefore $word.Documents.Count "$Name closed or created a source document."
    Assert-SourceUnchanged $before $Document $Name
    Complete-Scenario $Name
    if (-not $KeepOpen) { $Document.Close(0) }
    return $result
}

function New-DocxFixture {
    param([string]$Name, [switch]$Dirty)
    $path = Join-Path $fixtureRoot $Name
    $macroEnabled = [IO.Path]::GetExtension($Name) -ieq '.docm'
    $mainType = if ($macroEnabled) {
        'application/vnd.ms-word.document.macroEnabled.main+xml'
    }
    else {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
    }
    $documentBody = if ($Dirty) {
        '<w:commentRangeStart w:id="0"/><w:ins w:id="1" w:author="Fixture" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Tracked insertion</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> baseline</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
    }
    else {
        '<w:r><w:t>Clean fixture.</w:t></w:r>'
    }
    $contentTypes = @'
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="MAIN_TYPE"/>CONTENT_OVERRIDES</Types>
'@.Replace('MAIN_TYPE', $mainType)
    $rootRelationships = @'
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
'@
    $documentRelationships = ''
    $settings = ''
    $comments = ''
    if ($Dirty) {
        $contentTypes = $contentTypes.Replace('CONTENT_OVERRIDES', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>')
        $documentRelationships = @'
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>
'@
        $settings = '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:trackRevisions/></w:settings>'
        $comments = '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Fixture" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Synthetic comment</w:t></w:r></w:p></w:comment></w:comments>'
    }
    else {
        $contentTypes = $contentTypes.Replace('CONTENT_OVERRIDES', '')
    }
    $document = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>' + $documentBody + '</w:p><w:sectPr/></w:body></w:document>'
    $zip = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $entries = [ordered]@{
            '[Content_Types].xml' = $contentTypes
            '_rels/.rels' = $rootRelationships
            'word/document.xml' = $document
        }
        if ($Dirty) {
            $entries['word/_rels/document.xml.rels'] = $documentRelationships
            $entries['word/settings.xml'] = $settings
            $entries['word/comments.xml'] = $comments
        }
        foreach ($entryName in $entries.Keys) {
            $entry = $zip.CreateEntry($entryName)
            $stream = $entry.Open()
            $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
            try { $writer.Write($entries[$entryName]) } finally { $writer.Dispose() }
        }
    }
    finally { $zip.Dispose() }
    return $path
}

function Get-CellText {
    param($Table, [int]$Row, [int]$Column)
    return ([string]$Table.Cell($Row, $Column).Range.Text).Trim([char[]]@([char]13, [char]7))
}

try {
    $null = New-Item -ItemType Directory -Path $fixtureRoot
    $word = Start-RAToolsWordSession
    if ($word.Documents.Count -ne 0) {
        throw 'The new Word test session unexpectedly contains existing documents.'
    }
    $ownsSession = $true
    $word.AddIns.Unload($false)
    $hostDoc = $word.Documents.Add()
    $hostDoc.Content.Text = 'Clean detection test host.'
    foreach ($sourceFile in (Get-RAToolsVbaSourceFiles -RepoRoot $repoRoot)) {
        $null = $hostDoc.VBProject.VBComponents.Import($sourceFile.Path)
    }
    $hostDoc.Variables.Add('RAToolsCleanProbeArgument', 'False') | Out-Null
    foreach ($variable in @('RAToolsCleanFileCompleted', 'RAToolsCleanOpenedReadOnly', 'RAToolsCleanOpenedSecurity', 'RAToolsCleanWorkerWindow')) {
        $hostDoc.Variables.Add($variable, '0') | Out-Null
    }
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class RAToolsCleanTestWindows {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hWnd);
}
'@

    $observer = $hostDoc.VBProject.VBComponents.Add(2)
    $observer.Name = 'RAToolsCleanOpenObserver'
    $observer.CodeModule.AddFromString(@'
Public WithEvents App As Word.Application
Public OpenCount As Long
Public AllReadOnly As Boolean
Public ObservedSecurity As Long
Public CancelOwnedClose As Boolean

Private Sub App_DocumentOpen(ByVal Doc As Document)
    OpenCount = OpenCount + 1
    If Not Doc.ReadOnly Then AllReadOnly = False
    ObservedSecurity = App.AutomationSecurity
End Sub

Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    If CancelOwnedClose And Doc.ReadOnly Then Cancel = True
End Sub
'@)

    $component = $hostDoc.VBProject.VBComponents.Item('Macro_BatchDetectCleanVersions')
    # Observe the real open result without replacing the production open/check/close path.
    for ($line = 1; $line -le $component.CodeModule.CountOfLines; $line++) {
        if ($component.CodeModule.Lines($line, 1).Trim() -eq 'If doc Is Nothing Or Len(openError) > 0 Then') {
            $component.CodeModule.InsertLines($line, @'
    If Not doc Is Nothing Then
        ThisDocument.Variables("RAToolsCleanOpenedReadOnly").Value = CStr(doc.ReadOnly)
        ThisDocument.Variables("RAToolsCleanOpenedSecurity").Value = CStr(doc.Application.AutomationSecurity)
        If Not inspectionApp Is Nothing Then ThisDocument.Variables("RAToolsCleanWorkerWindow").Value = CStr(doc.Windows(1).Hwnd)
    End If
'@)
            break
        }
    }
    $component.CodeModule.InsertLines(2, 'Private cleanObserver As RAToolsCleanOpenObserver')
    $component.CodeModule.AddFromString(@'
Public Function RAToolsCleanDocumentProbe() As String
    Dim result As Variant
    result = InspectCleanDocument(Application.Documents(ThisDocument.Variables("RAToolsCleanProbeArgument").Value))
    RAToolsCleanDocumentProbe = result(0) & vbTab & result(1)
End Function

Public Function RAToolsCleanFileProbe() As String
    Dim result As Variant
    result = InspectCleanFile(ThisDocument.Variables("RAToolsCleanProbeArgument").Value)
    ThisDocument.Variables("RAToolsCleanFileCompleted").Value = "True"
    RAToolsCleanFileProbe = result(1) & vbTab & result(2)
End Function

Public Sub RAToolsCleanCreateFileFixtures()
    Dim fixture As Document
    Dim fixtureFolder As String
    Dim eventCode As String
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    fixtureFolder = ThisDocument.Variables("RAToolsCleanProbeArgument").Value
    Set fixture = Application.Documents.Add
    fixture.TrackRevisions = False
    fixture.Content.Text = "Binary DOC fixture."
    fixture.SaveAs2 FileName:=fixtureFolder & "\legacy-clean.doc", _
                    FileFormat:=wdFormatDocument97, AddToRecentFiles:=False
    fixture.TrackRevisions = True
    fixture.Content.InsertAfter " Tracked binary DOC insertion."
    fixture.TrackRevisions = False
    fixture.SaveAs2 FileName:=fixtureFolder & "\legacy-dirty.doc", _
                    FileFormat:=wdFormatDocument97, AddToRecentFiles:=False
    fixture.Close SaveChanges:=wdDoNotSaveChanges
    Set fixture = Nothing

    Set fixture = Application.Documents.Add
    fixture.TrackRevisions = False
    fixture.Content.Text = "Document_Open fixture."
    eventCode = "Private Sub Document_Open()" & vbCrLf & _
        "    Dim marker As Integer" & vbCrLf & _
        "    marker = FreeFile" & vbCrLf & _
        "    Open """ & fixtureFolder & "\open-event.txt"" For Append As #marker" & vbCrLf & _
        "    Print #marker, ""Document_Open executed""" & vbCrLf & _
        "    Close #marker" & vbCrLf & _
        "    Me.TrackRevisions = True" & vbCrLf & _
        "    Me.Content.InsertAfter "" Document_Open executed.""" & vbCrLf & _
        "End Sub"
    fixture.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString eventCode
    fixture.SaveAs2 FileName:=fixtureFolder & "\event-clean.docm", _
                    FileFormat:=wdFormatXMLDocumentMacroEnabled, AddToRecentFiles:=False
    fixture.Close SaveChanges:=wdDoNotSaveChanges
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise failureNumber, "Clean detection fixture creation", failureDescription
End Sub

Public Function RAToolsCleanBatchProbe() As String
    Dim fileList As New Collection
    Dim path As Variant
    Dim report As Document
    For Each path In Split(ThisDocument.Variables("RAToolsCleanProbeArgument").Value, vbLf)
        fileList.Add CStr(path)
    Next path
    Set report = DetectCleanFiles(fileList)
    RAToolsCleanBatchProbe = report.Name
End Function

Public Function RAToolsCleanMissingProbe() As String
    Dim missing As Document
    Dim result As Variant
    result = InspectCleanDocument(missing)
    RAToolsCleanMissingProbe = result(0) & vbTab & result(1)
End Function

Public Function RAToolsCleanClosedProbe() As String
    Dim retired As Document
    Dim result As Variant
    Set retired = Application.Documents.Add
    retired.Close SaveChanges:=wdDoNotSaveChanges
    result = InspectCleanDocument(retired)
    RAToolsCleanClosedProbe = result(0) & vbTab & result(1)
End Function

Public Sub RAToolsCleanStartObserver()
    If Not cleanObserver Is Nothing Then Set cleanObserver.App = Nothing
    Set cleanObserver = New RAToolsCleanOpenObserver
    cleanObserver.AllReadOnly = True
    cleanObserver.CancelOwnedClose = (ThisDocument.Variables("RAToolsCleanProbeArgument").Value = "True")
    Set cleanObserver.App = Application
End Sub

Public Function RAToolsCleanObserverProbe() As String
    RAToolsCleanObserverProbe = CStr(cleanObserver.OpenCount) & vbTab & CStr(cleanObserver.AllReadOnly) & vbTab & CStr(cleanObserver.ObservedSecurity)
End Function

Public Sub RAToolsCleanStopObserver()
    If Not cleanObserver Is Nothing Then Set cleanObserver.App = Nothing
    Set cleanObserver = Nothing
End Sub

Public Function RAToolsCleanMacroListProbe() As String
    Dim menu As New frmMacroList
    Dim row As Long
    Load menu
    menu.txtSearch.Text = "BatchDetectCleanVersions"
    For row = 0 To menu.lstMacros.ListCount - 1
        If menu.lstMacros.List(row, 0) = "BatchDetectCleanVersions" Then
            RAToolsCleanMacroListProbe = menu.lstMacros.List(row, 1) & vbTab & menu.lstMacros.List(row, 2)
            Exit For
        End If
    Next row
    Unload menu
End Function
'@)

    $source = New-SourceFixture
    $field = $source.Fields.Add($source.Range(0, 0), 33)
    $field.Locked = $true
    $null = Assert-Classification $source 'Clean' @('无修订', '无批注', '修订跟踪已关闭') 'Clean with locked fields' -KeepOpen
    Assert-Equal $true ([bool]$field.Locked) 'Detection unlocked a field.'
    $source.Close(0)

    $source = New-SourceFixture
    $source.Content.HighlightColorIndex = 7
    $source.Content.Font.Hidden = -1
    $source.Variables.Add('IgnoredFixtureMetadata', 'Synthetic fixture') | Out-Null
    $null = Assert-Classification $source 'Clean' @() 'Highlight, hidden text and metadata are outside the definition'

    $source = New-SourceFixture
    $source.TrackRevisions = $true
    $result = Assert-Classification $source '非 Clean' @('修订跟踪仍开启') 'Tracking enabled without revisions'
    Assert-True (-not $result.Reason.Contains('存在修订')) 'Tracking alone was reported as a revision.'
    Assert-True (-not $result.Reason.Contains('存在批注')) 'Tracking alone was reported as a comment.'

    $source = New-SourceFixture
    $source.TrackRevisions = $true
    $source.Range(0, 0).InsertBefore('Tracked insertion. ')
    $source.TrackRevisions = $false
    Assert-True ($source.Revisions.Count -gt 0) 'Revision fixture has no revision.'
    $result = Assert-Classification $source '非 Clean' @('存在修订') 'Revisions with tracking turned off'
    Assert-True (-not $result.Reason.Contains('修订跟踪仍开启')) 'Closed tracking was reported as enabled.'

    $source = New-SourceFixture
    $source.Comments.Add($source.Range(0, 6), 'Synthetic comment') | Out-Null
    $null = Assert-Classification $source '非 Clean' @('存在批注') 'Comments without revisions'

    $source = New-SourceFixture
    $source.TrackRevisions = $true
    $source.Range(0, 0).InsertBefore('Tracked insertion. ')
    $source.Comments.Add($source.Range(0, 6), 'Synthetic comment') | Out-Null
    $null = Assert-Classification $source '非 Clean' @('存在修订', '存在批注', '修订跟踪仍开启') 'All failed conditions are listed'

    $source = New-SourceFixture
    $source.Protect(3)
    Assert-Equal 3 ([int]$source.ProtectionType) 'Protected fixture is not protected.'
    $null = Assert-Classification $source 'Clean' @() 'Readable protected document is not unprotected'

    $stories = @(
        @{ Name = 'Footnote'; Setup = { param($d) ($d.Footnotes.Add($d.Range(1, 1), '', 'Footnote baseline')).Range } },
        @{ Name = 'Endnote'; Setup = { param($d) ($d.Endnotes.Add($d.Range(1, 1), '', 'Endnote baseline')).Range } },
        @{ Name = 'Second section header'; Setup = {
            param($d)
            $d.Range($d.Content.End - 1, $d.Content.End - 1).InsertBreak(2)
            $header = $d.Sections.Item(2).Headers.Item(1)
            $header.LinkToPrevious = $false
            $header.Range.Text = 'Second header baseline'
            $header.Range
        } },
        @{ Name = 'First page header'; Setup = {
            param($d)
            $d.PageSetup.DifferentFirstPageHeaderFooter = -1
            $header = $d.Sections.Item(1).Headers.Item(2)
            $header.Range.Text = 'First header baseline'
            $header.Range
        } },
        @{ Name = 'Even page footer'; Setup = {
            param($d)
            $d.PageSetup.OddAndEvenPagesHeaderFooter = -1
            $footer = $d.Sections.Item(1).Footers.Item(3)
            $footer.Range.Text = 'Even footer baseline'
            $footer.Range
        } },
        @{ Name = 'Text box'; Setup = {
            param($d)
            $shape = $d.Shapes.AddTextbox(1, 10, 10, 120, 50, $d.Range(1, 1))
            $shape.TextFrame.TextRange.Text = 'Text box baseline'
            $shape.TextFrame.TextRange
        } },
        @{ Name = 'Canvas text box'; Setup = {
            param($d)
            $canvas = $d.Shapes.AddCanvas(10, 10, 220, 100, $d.Range(1, 1))
            $box = $canvas.CanvasItems.AddTextbox(1, 0, 0, 120, 50)
            $box.TextFrame.TextRange.Text = 'Canvas baseline'
            $box.TextFrame.TextRange
        } }
    )
    foreach ($story in $stories) {
        $source = New-SourceFixture
        $range = & $story.Setup $source
        $source.TrackRevisions = $true
        $range.InsertBefore('Tracked story insertion. ')
        $source.TrackRevisions = $false
        Assert-True ($range.Revisions.Count -gt 0) "$($story.Name) fixture has no revision."
        Assert-Equal 0 $source.Revisions.Count "$($story.Name) fixture also has main-story revisions."
        $null = Assert-Classification $source '非 Clean' @('存在修订') "$($story.Name) revisions"
    }

    $source = New-SourceFixture
    $source.Shapes.AddLine(10, 10, 50, 50, $source.Range(1, 1)) | Out-Null
    $source.Shapes.AddTextbox(1, 70, 10, 120, 50, $source.Range(1, 1)) | Out-Null
    $source.Frames.Add($source.Range(0, 6)) | Out-Null
    $null = Assert-Classification $source 'Clean' @() 'Line, empty text box and frame do not cause false errors'

    foreach ($probe in @('RAToolsCleanMissingProbe', 'RAToolsCleanClosedProbe')) {
        $result = ConvertFrom-ProbeResult ([string](Invoke-CleanProbe $probe))
        Assert-Equal '无法检测' $result.Status "$probe must not report Clean."
        Assert-True ($result.Reason.Contains('检查失败')) "$probe should include the checking error."
        Complete-Scenario $probe
    }

    $cleanPath = New-DocxFixture 'clean.docx'
    Write-Host 'Creating real DOC and Document_Open DOCM fixtures in Word.'
    $null = Invoke-CleanProbe 'RAToolsCleanCreateFileFixtures' $fixtureRoot
    $legacyPath = Join-Path $fixtureRoot 'legacy-clean.doc'
    $legacyDirtyPath = Join-Path $fixtureRoot 'legacy-dirty.doc'
    $macroPath = Join-Path $fixtureRoot 'event-clean.docm'
    $markerPath = Join-Path $fixtureRoot 'open-event.txt'
    foreach ($path in @($legacyPath, $legacyDirtyPath)) {
        $signature = [BitConverter]::ToString([IO.File]::ReadAllBytes($path), 0, 8)
        Assert-Equal 'D0-CF-11-E0-A1-B1-1A-E1' $signature 'DOC fixture must be a binary compound file.'
    }
    $macroPackage = [IO.Compression.ZipFile]::OpenRead($macroPath)
    try { Assert-True ($null -ne $macroPackage.GetEntry('word/vbaProject.bin')) 'DOCM fixture has no VBA project.' }
    finally { $macroPackage.Dispose() }
    Assert-Equal $false (Test-Path -LiteralPath $markerPath) 'Opening event ran while making the fixture.'
    $macroHashBefore = (Get-FileHash -LiteralPath $macroPath -Algorithm SHA256).Hash
    $openPath = New-DocxFixture 'already-open.docx'
    $dirtyPath = New-DocxFixture 'dirty.docx' -Dirty
    $missingPath = Join-Path $fixtureRoot 'missing.docx'
    $brokenPath = Join-Path $fixtureRoot 'broken.docx'
    [IO.File]::WriteAllBytes($brokenPath, [byte[]]@(0x50, 0x4B, 0x03, 0x04, 0, 0, 0, 0))

    $null = Invoke-CleanProbe 'RAToolsCleanStartObserver'
    foreach ($path in @($cleanPath, $legacyPath, $macroPath)) {
        $hashBefore = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $countBefore = $word.Documents.Count
        $securityBefore = $word.AutomationSecurity
        $result = Invoke-FileCheck $path
        Assert-Equal 'Clean' $result.Status "Read-only file detection for $([IO.Path]::GetExtension($path))."
        Assert-Equal $countBefore $word.Documents.Count 'A tool-opened source was left open.'
        Assert-Equal $securityBefore $word.AutomationSecurity 'AutomationSecurity was not restored.'
        Assert-Equal 'True' ([string]$hostDoc.Variables.Item('RAToolsCleanOpenedReadOnly').Value) 'A source was opened without ReadOnly.'
        Assert-Equal 3 ([int]$hostDoc.Variables.Item('RAToolsCleanOpenedSecurity').Value) 'Source macros were not disabled during file open.'
        Assert-Equal $hashBefore (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash 'Source file bytes changed.'
    }
    $observed = ([string](Invoke-CleanProbe 'RAToolsCleanObserverProbe')) -split ([char]9)
    Assert-Equal 1 ([int]$observed[0]) 'Only DOCX should open in the calling Word instance.'
    Assert-Equal 'True' $observed[1] 'A source was opened without ReadOnly.'
    Assert-Equal 3 ([int]$observed[2]) 'Source macros were not disabled during file open.'
    Assert-Equal $false (Test-Path -LiteralPath $markerPath) 'Detection executed Document_Open.'
    Complete-Scenario 'Real DOC, DOCX and event DOCM are read-only with Document_Open blocked'

    $securityBefore = $word.AutomationSecurity
    try {
        $word.AutomationSecurity = 1
        $eventSource = $word.Documents.Open($macroPath, $false, $true)
        try {
            Assert-Equal $true (Test-Path -LiteralPath $markerPath) 'Positive control did not execute Document_Open.'
            Assert-Equal 1 ([IO.File]::ReadAllLines($markerPath).Count) 'Positive control did not record exactly one event.'
            Assert-Equal $true ([bool]$eventSource.TrackRevisions) 'Positive control did not enable revision tracking.'
            Assert-True ($eventSource.Revisions.Count -gt 0) 'Positive control did not create a revision.'
        }
        finally { $eventSource.Close(0) }
    }
    finally { $word.AutomationSecurity = $securityBefore }
    Assert-Equal $macroHashBefore (Get-FileHash -LiteralPath $macroPath -Algorithm SHA256).Hash 'Positive control changed the DOCM file.'
    Complete-Scenario 'Real Document_Open executes in the enabled-macro positive control'

    $hashBefore = (Get-FileHash -LiteralPath $legacyDirtyPath -Algorithm SHA256).Hash
    $countBefore = $word.Documents.Count
    $result = Invoke-FileCheck $legacyDirtyPath
    Assert-Equal '非 Clean' $result.Status 'Binary DOC revisions were not detected.'
    Assert-True ($result.Reason.Contains('存在修订')) 'Binary DOC has no revision reason.'
    Assert-Equal $countBefore $word.Documents.Count 'Binary DOC was left open.'
    Assert-Equal $hashBefore (Get-FileHash -LiteralPath $legacyDirtyPath -Algorithm SHA256).Hash 'Binary DOC bytes changed.'
    Complete-Scenario 'Real binary DOC revisions are detected without saving the file'

    $openHashBefore = (Get-FileHash -LiteralPath $openPath -Algorithm SHA256).Hash
    $existing = $word.Documents.Open($openPath, $false, $false)
    $existing.TrackRevisions = $true
    $existing.Content.InsertAfter(' Unsaved tracked edit.')
    Assert-Equal $false ([bool]$existing.Saved) 'Existing fixture has no unsaved edits.'
    $before = Get-SourceSnapshot $existing
    $countBefore = $word.Documents.Count
    $null = Invoke-CleanProbe 'RAToolsCleanStartObserver'
    $result = Invoke-FileCheck $openPath
    Assert-Equal '非 Clean' $result.Status 'Already-open detection ignored unsaved revisions.'
    Assert-Equal $countBefore $word.Documents.Count 'Already-open document was replaced or closed.'
    Assert-SourceUnchanged $before $existing 'Already-open detection'
    $observed = ([string](Invoke-CleanProbe 'RAToolsCleanObserverProbe')) -split ([char]9)
    Assert-Equal 0 ([int]$observed[0]) 'An already-open document was reopened.'
    Complete-Scenario 'Already-open unsaved document is reused and preserved'

    foreach ($path in @($missingPath, $brokenPath)) {
        $result = Invoke-FileCheck $path
        Assert-Equal '无法检测' $result.Status 'Opening failures must not be Clean.'
        Assert-True ($result.Reason.Contains('无法打开文档')) 'Opening failure has no opening error.'
    }
    Complete-Scenario 'Missing and corrupt files are unable to detect'

    $paths = [string[]]@($cleanPath, $missingPath, $dirtyPath, $openPath)
    $hashesBefore = @{}
    foreach ($path in @($cleanPath, $dirtyPath)) {
        $hashesBefore[$path] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
    $before = Get-SourceSnapshot $existing
    $countBefore = $word.Documents.Count
    $updatingBefore = $word.ScreenUpdating
    $alertsBefore = $word.DisplayAlerts
    $null = Invoke-CleanProbe 'RAToolsCleanStartObserver'
    $batchPaths = [string]::Join([char]10, $paths)
    $reportName = [string](Invoke-CleanProbe 'RAToolsCleanBatchProbe' $batchPaths)
    $report = $word.Documents.Item($reportName)
    $table = $report.Tables.Item(1)
    Assert-Equal ($countBefore + 1) $word.Documents.Count 'Batch should retain only its new report.'
    Assert-Equal 5 $table.Rows.Count 'Report should contain a heading and four results.'
    Assert-Equal 3 $table.Columns.Count 'Report columns should be path, status and reason.'
    Assert-Equal '文件路径' (Get-CellText $table 1 1) 'Report path heading.'
    Assert-Equal '结论' (Get-CellText $table 1 2) 'Report status heading.'
    Assert-Equal '原因' (Get-CellText $table 1 3) 'Report reason heading.'
    $expectedStatuses = @('Clean', '无法检测', '非 Clean', '非 Clean')
    for ($index = 0; $index -lt $paths.Count; $index++) {
        Assert-Equal $paths[$index] (Get-CellText $table ($index + 2) 1) 'Report path or ordering.'
        Assert-Equal $expectedStatuses[$index] (Get-CellText $table ($index + 2) 2) 'Report status or error isolation.'
    }
    $dirtyReason = Get-CellText $table 4 3
    foreach ($reason in @('存在修订', '存在批注', '修订跟踪仍开启')) {
        Assert-True ($dirtyReason.Contains($reason)) "Report omitted <$reason>."
    }
    $reportText = [string]$report.Content.Text
    Assert-True ($reportText.Contains('共 4 个文件；Clean 1；非 Clean 2；无法检测 1')) 'Report totals are incorrect.'
    Assert-Equal $false ([bool]$report.Saved) 'The report should remain open and unsaved.'
    Assert-Equal $false ([bool]$report.TrackRevisions) 'The report should not accumulate revisions.'
    Assert-Equal $updatingBefore $word.ScreenUpdating 'Batch changed ScreenUpdating.'
    Assert-Equal $alertsBefore $word.DisplayAlerts 'Batch changed DisplayAlerts.'
    Assert-SourceUnchanged $before $existing 'Batch detection'
    foreach ($path in @($cleanPath, $dirtyPath)) {
        Assert-Equal $hashesBefore[$path] (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash 'Batch changed source bytes.'
    }
    $observed = ([string](Invoke-CleanProbe 'RAToolsCleanObserverProbe')) -split ([char]9)
    Assert-Equal 2 ([int]$observed[0]) 'Batch should open only new readable sources.'
    Assert-Equal 'True' $observed[1] 'Batch opened a source without ReadOnly.'
    Complete-Scenario 'Mixed batch continues after errors and reports exact reasons and totals'

    $countBefore = $word.Documents.Count
    $null = Invoke-CleanProbe 'RAToolsCleanStartObserver' 'True'
    $result = Invoke-FileCheck $cleanPath
    Assert-Equal '无法检测' $result.Status 'Cancelled source close was reported as Clean.'
    Assert-True ($result.Reason.Contains('关闭失败')) 'Cancelled close has no reason.'
    Assert-Equal ($countBefore + 1) $word.Documents.Count 'Cancelled source close should retain its read-only document.'
    $null = Invoke-CleanProbe 'RAToolsCleanStopObserver'
    $word.Documents.Item($cleanPath).Close(0)
    Complete-Scenario 'Cancelled source closure is reported instead of silently passing'

    $hostDoc.Activate()
    $menu = ([string](Invoke-CleanProbe 'RAToolsCleanMacroListProbe')) -split ([char]9), 2
    Assert-Equal 2 $menu.Count 'Actual macro list contains no clean detection entry.'
    Assert-Equal '批量检测 clean 版本' $menu[0] 'Actual macro list has the wrong label.'
    Assert-True ($menu[1].Contains('不修改或保存源文档')) 'Actual macro list has no read-only description.'
    Complete-Scenario 'Real Word macro list exposes and searches the new entry'
    $existing.Close(0)
    Assert-Equal $openHashBefore (Get-FileHash -LiteralPath $openPath -Algorithm SHA256).Hash 'Unsaved edits were written to disk.'

    Write-Host "PASS BatchDetectCleanVersions.Tests ($script:scenarioCount scenarios; $script:assertionCount assertions)"
}
finally {
    if ($ownsSession) {
        try { $null = Invoke-CleanProbe 'RAToolsCleanStopObserver' } catch { }
        foreach ($document in @($word.Documents)) {
            $document.Close(0)
        }
        Stop-RAToolsWordSession -Word $word
    }
    elseif ($null -ne $word -and [Runtime.InteropServices.Marshal]::IsComObject($word)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    if (Test-Path -LiteralPath $fixtureRoot) {
        Move-RAToolsToRecycleBin -Path $fixtureRoot -SafeRoot $repoRoot
    }
}
