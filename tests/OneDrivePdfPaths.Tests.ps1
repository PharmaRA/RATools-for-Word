#requires -version 5.1
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'scripts\RATools.Build.psm1') -Force
Add-Type -AssemblyName Microsoft.VisualBasic

# Exercise production VBA with synthetic cloud metadata and mappings.
# Word still opens local fixtures and performs real PDF exports.
$fixtureRoot = Join-Path (Join-Path $repoRoot 'backups') ('RATools_OneDrivePdf_' + [guid]::NewGuid().ToString('N'))
$businessRoot = Join-Path $fixtureRoot 'Business synced'
$personalRoot = Join-Path $fixtureRoot 'Personal synced'
$businessFolder = Join-Path $businessRoot '子文件夹'
$personalFolder = Join-Path $personalRoot 'Quarter Reports'
$chosenFolder = Join-Path $fixtureRoot 'Chosen output'
foreach ($folder in @($businessFolder,$personalFolder,$chosenFolder)) {
    [void][IO.Directory]::CreateDirectory($folder)
}
$gbk = [Text.Encoding]::GetEncoding(936)
$word = $null
$script:assertionCount = 0

function Assert-True([bool]$Condition,[string]$Message) {
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
    $script:assertionCount++
}
function Assert-Pdf([string]$Path) {
    Assert-True (Test-Path -LiteralPath $Path -PathType Leaf) 'PDF output must exist.'
    $bytes = [IO.File]::ReadAllBytes($Path)
    Assert-True ($bytes.Length -gt 5 -and [Text.Encoding]::ASCII.GetString($bytes,0,5) -eq '%PDF-') 'Output must contain real PDF bytes.'
}
function Run-Probe([string]$Name,[object[]]$Arguments) {
    $hostDoc.Activate()
    $invokeArgs = @($Name) + $Arguments
    return $word.GetType().InvokeMember('Run',[Reflection.BindingFlags]::InvokeMethod,$null,$word,$invokeArgs)
}
function Replace-OneBlock([string]$Text,[string]$Pattern,[string]$Replacement) {
    $regex = [regex]::new($Pattern)
    if ($regex.Matches($Text).Count -ne 1) { throw 'Expected exactly one VBA test seam.' }
    return $regex.Replace($Text,[Text.RegularExpressions.MatchEvaluator]{ param($match) $Replacement })
}

try {
    $word = Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $hostDoc = $word.Documents.Add()
    $core = [IO.File]::ReadAllText((Join-Path $repoRoot 'modules\Core_Files.bas'),$gbk)
    $mappingStub = @"
Private Sub LoadOneDriveSyncMappings(ByRef urlNamespaces As Collection, ByRef localRoots As Collection)
    urlNamespaces.Add "https://example.invalid/sites/ra/Shared%20Documents"
    localRoots.Add "$businessRoot"
    urlNamespaces.Add "https://d.docs.live.net/0123456789ABCDEF"
    localRoots.Add "$personalRoot"
End Sub
"@
    $core = Replace-OneBlock $core '(?ms)^Private Sub LoadOneDriveSyncMappings\(.*?^End Sub' $mappingStub
    $corePath = Join-Path $fixtureRoot 'Core_Files.bas'
    [IO.File]::WriteAllText($corePath,$core,$gbk)
    [void]$hostDoc.VBProject.VBComponents.Import($corePath)
    foreach ($name in @('Core_Session','Core_UI')) {
        [void]$hostDoc.VBProject.VBComponents.Import((Join-Path $repoRoot "modules\$name.bas"))
    }

    $macro = [IO.File]::ReadAllText((Join-Path $repoRoot 'modules\Macro_BatchConvertWordToPDF.bas'),$gbk)
    $declarations = @'
Option Explicit
Private probeDocumentFolder As String
Private probeSelectedFolder As String
Private probePromptCount As Long
Private probeLocalSource As String
'@
    $macro = $macro.Replace('Option Explicit',$declarations)
    $macro = [regex]::Replace($macro,'\bMsgBox\b','PdfProbeMessage')
    $macro = $macro.Replace('doc.Path','PdfProbeDocumentFolder(doc)')
    $macro = $macro.Replace('PickFolder(','PdfProbePickFolder(')
    $openStatement = 'Set doc = Documents.Open(FileName:=filePath, Visible:=True, ReadOnly:=True, AddToRecentFiles:=False)'
    Assert-True ($macro.Contains($openStatement)) 'Batch test must intercept only document opening.'
    $macro = $macro.Replace($openStatement,'Set doc = PdfProbeOpenDocument(filePath)')
    $macro += @'

Private Function PdfProbeMessage(ByVal text As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
End Function
Private Function PdfProbeDocumentFolder(ByVal doc As Document) As String
    If Len(probeDocumentFolder) > 0 Then
        PdfProbeDocumentFolder = probeDocumentFolder
    Else
        PdfProbeDocumentFolder = doc.Path
    End If
End Function
Private Function PdfProbePickFolder(ByVal title As String) As String
    probePromptCount = probePromptCount + 1
    PdfProbePickFolder = probeSelectedFolder
End Function
Private Function PdfProbeOpenDocument(ByVal pathValue As String) As Document
    If IsWebPath(pathValue) Then pathValue = probeLocalSource
    Set PdfProbeOpenDocument = Documents.Open(FileName:=pathValue, Visible:=True, ReadOnly:=True, AddToRecentFiles:=False)
End Function
Public Function PdfProbeBuild(ByVal sourceFolder As String, ByVal documentName As String, ByVal selectedFolder As String) As String
    Dim session As ConvertSession, output As String
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    probePromptCount = 0
    probeSelectedFolder = selectedFolder
    On Error GoTo Failed
    output = BuildPdfOutputPath(session, sourceFolder, documentName)
    PdfProbeBuild = output & "|" & CStr(probePromptCount)
    Exit Function
Failed:
    PdfProbeBuild = "ERROR:" & CStr(Err.Number) & "|" & CStr(probePromptCount)
End Function
Public Function PdfProbeReuse(ByVal selectedFolder As String) As String
    Dim session As ConvertSession, firstPath As String, secondPath As String
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    probePromptCount = 0
    probeSelectedFolder = selectedFolder
    firstPath = BuildPdfOutputPath(session, "https://unmapped.invalid/library", "First.docx")
    secondPath = BuildPdfOutputPath(session, "https://another.invalid/library", "Second.docm")
    PdfProbeReuse = firstPath & "|" & secondPath & "|" & CStr(probePromptCount)
End Function
Public Function PdfProbeCancelTwice() As String
    Dim session As ConvertSession, output As String, firstError As Long, secondError As Long
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    probePromptCount = 0
    probeSelectedFolder = ""
    On Error Resume Next
    output = BuildPdfOutputPath(session, "https://unmapped.invalid/library", "First.docx")
    firstError = Err.Number
    Err.Clear
    output = BuildPdfOutputPath(session, "https://unmapped.invalid/library", "Second.docx")
    secondError = Err.Number
    On Error GoTo 0
    PdfProbeCancelTwice = CStr(firstError) & "|" & CStr(secondError) & "|" & CStr(probePromptCount)
End Function
Public Function PdfProbeActive(ByVal target As Document, ByVal reportedFolder As String, ByVal selectedFolder As String) As String
    Dim session As ConvertSession
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    session.updateType = 1
    probeDocumentFolder = reportedFolder
    probeSelectedFolder = selectedFolder
    probePromptCount = 0
    target.Activate
    ConvertActiveDocument session
    PdfProbeActive = CStr(session.successCount) & "|" & CStr(session.failCount) & "|" & CStr(probePromptCount) & "|" & session.lastOutputPath
End Function
Public Function PdfProbeBatch(ByVal source As String, ByVal reportedFolder As String, ByVal localSource As String, ByVal selectedFolder As String) As String
    Dim session As ConvertSession
    Set session.fso = CreateObject("Scripting.FileSystemObject")
    session.updateType = 1
    session.totalFileCount = 1
    probeDocumentFolder = reportedFolder
    probeSelectedFolder = selectedFolder
    probeLocalSource = localSource
    probePromptCount = 0
    ConvertOneFile session, source
    PdfProbeBatch = CStr(session.successCount) & "|" & CStr(session.failCount) & "|" & CStr(probePromptCount) & "|" & session.lastOutputPath
End Function
'@
    $macro = $macro.Replace("`r`n","`n").Replace("`n","`r`n")
    $macroPath = Join-Path $fixtureRoot 'Macro_BatchConvertWordToPDF.bas'
    [IO.File]::WriteAllText($macroPath,$macro,$gbk)
    [void]$hostDoc.VBProject.VBComponents.Import($macroPath)
    $hostDoc.Activate()

    Write-Host 'Running synthetic OneDrive path probes'
    $cloudFolder = 'https://example.invalid/sites/ra/Shared%20Documents/%E5%AD%90%E6%96%87%E4%BB%B6%E5%A4%B9?download=1#section'
    $business = [string](Run-Probe 'PdfProbeBuild' @($cloudFolder,'Cloud report.docx',''))
    Assert-True ($business -eq ((Join-Path $businessFolder 'Cloud report.pdf') + '|0')) 'Business OneDrive paths must decode UTF-8 names and use the synced folder.'
    $personal = [string](Run-Probe 'PdfProbeBuild' @('https://d.docs.live.net/0123456789abcdef/Quarter%20Reports','Report+notes.docm',''))
    Assert-True ($personal -eq ((Join-Path $personalFolder 'Report+notes.pdf') + '|0')) 'Personal OneDrive paths must resolve without changing literal plus signs.'
    $local = [string](Run-Probe 'PdfProbeBuild' @($fixtureRoot,'Local.doc',''))
    Assert-True ($local -eq ((Join-Path $fixtureRoot 'Local.pdf') + '|0')) 'Local folders must retain the original output location.'
    $reuse = [string](Run-Probe 'PdfProbeReuse' @($chosenFolder))
    Assert-True ($reuse -eq ((Join-Path $chosenFolder 'First.pdf') + '|' + (Join-Path $chosenFolder 'Second.pdf') + '|1')) 'Unmapped cloud documents must share one user-selected output folder.'
    $cancelError = -2147221504 + 2401
    $cancelTwice = [string](Run-Probe 'PdfProbeCancelTwice' @())
    Assert-True ($cancelTwice -eq "$cancelError|$cancelError|1") 'Canceling must give an explicit error without repeating the folder prompt.'
    $invalid = [string](Run-Probe 'PdfProbeBuild' @('https://unmapped.invalid/library','Invalid.docx','https://unmapped.invalid/output'))
    Assert-True ($invalid -eq ('ERROR:' + (-2147221504 + 2402) + '|1')) 'A web URL must be rejected before any filesystem output operation.'
    $missing = [string](Run-Probe 'PdfProbeBuild' @('https://unmapped.invalid/library','Missing.docx',(Join-Path $fixtureRoot 'Not created')))
    Assert-True ($missing -eq ('ERROR:' + (-2147221504 + 2403) + '|1')) 'An unavailable selected folder must give an explicit output-location error.'

    Write-Host 'Running real Word PDF export probes with synthetic cloud metadata'
    $cloudDoc = $word.Documents.Add()
    $cloudDoc.Content.Text = 'Synthetic OneDrive PDF fixture with pending edits.'
    $lockedField = $cloudDoc.Fields.Add($cloudDoc.Range(0,0),33,'PAGE')
    $lockedField.Locked = $true
    $unlockedField = $cloudDoc.Fields.Add($cloudDoc.Range(0,0),33,'PAGE')
    $unlockedField.Locked = $false
    $pendingText = $cloudDoc.Content.Text
    $savedBefore = $cloudDoc.Saved
    $countBefore = $word.Documents.Count
    $active = ([string](Run-Probe 'PdfProbeActive' @($cloudDoc,$cloudFolder,''))).Split('|')
    Assert-True ($active[0] -eq '1' -and $active[1] -eq '0' -and $active[2] -eq '0') 'The current cloud document must export to its synced folder.'
    Assert-True ((Split-Path -Parent $active[3]) -eq $businessFolder) 'Current-document output must use the mapped local folder.'
    Assert-Pdf $active[3]
    Assert-True ($lockedField.Locked -and -not $unlockedField.Locked) 'Export must restore mixed field locks.'
    Assert-True ($cloudDoc.Content.Text -eq $pendingText -and $cloudDoc.Saved -eq $savedBefore -and $word.Documents.Count -eq $countBefore) 'Export must retain the open document and pending edits.'
    $fallback = ([string](Run-Probe 'PdfProbeActive' @($cloudDoc,'https://unmapped.invalid/library',$chosenFolder))).Split('|')
    Assert-True ($fallback[0] -eq '1' -and $fallback[1] -eq '0' -and $fallback[2] -eq '1') 'An unmapped current document must export after choosing a local folder.'
    Assert-True ((Split-Path -Parent $fallback[3]) -eq $chosenFolder) 'Manual output must use the chosen folder.'
    Assert-Pdf $fallback[3]
    $cancelled = ([string](Run-Probe 'PdfProbeActive' @($cloudDoc,'https://unmapped.invalid/library',''))).Split('|')
    Assert-True ($cancelled[0] -eq '0' -and $cancelled[1] -eq '1' -and $cancelled[2] -eq '1') 'Canceling output selection must report conversion failure.'
    Assert-True ($lockedField.Locked -and -not $unlockedField.Locked -and $cloudDoc.Content.Text -eq $pendingText -and $cloudDoc.Saved -eq $savedBefore) 'Canceled conversion must retain field locks and pending edits.'

    $localSource = Join-Path $fixtureRoot 'LocalFixture.docx'
    $fixture = $word.Documents.Add()
    $fixture.Content.Text = 'Synthetic batch input.'
    $fixture.SaveAs2([string]$localSource,12)
    $fixture.Close(0)
    $hostDoc.Activate()
    $countBefore = $word.Documents.Count
    $localBatch = ([string](Run-Probe 'PdfProbeBatch' @($localSource,$cloudFolder,'',''))).Split('|')
    Assert-True ($localBatch[0] -eq '1' -and $localBatch[1] -eq '0' -and $localBatch[2] -eq '0') 'A locally selected cloud-synced file must export without asking for a folder.'
    Assert-True ($localBatch[3] -eq (Join-Path $fixtureRoot 'LocalFixture.pdf')) 'The selected local source path must take precedence over Word cloud metadata.'
    Assert-Pdf $localBatch[3]
    Assert-True ($word.Documents.Count -eq $countBefore) 'A document opened by the batch must close after export.'
    $urlSource = 'https://example.invalid/sites/ra/Shared%20Documents/LocalFixture.docx'
    $urlBatch = ([string](Run-Probe 'PdfProbeBatch' @($urlSource,$cloudFolder,$localSource,''))).Split('|')
    Assert-True ($urlBatch[0] -eq '1' -and $urlBatch[1] -eq '0' -and $urlBatch[2] -eq '0') 'A URL-selected document must export using its mapped document folder.'
    Assert-True ($urlBatch[3] -eq (Join-Path $businessFolder 'LocalFixture.pdf')) 'URL input must never become a filesystem output path.'
    Assert-Pdf $urlBatch[3]
    Assert-True ($word.Documents.Count -eq $countBefore) 'A URL-selected document opened by the batch must close after export.'
    Write-Host "PASS OneDrivePdfPaths.Tests ($script:assertionCount assertions)"
} finally {
    if ($null -ne $word) {
        try { for ($i = $word.Documents.Count; $i -ge 1; $i--) { $word.Documents.Item($i).Close(0) } } catch {}
    }
    Stop-RAToolsWordSession -Word $word
    $cleanupPath = [IO.Path]::GetFullPath($fixtureRoot)
    $allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'backups')) + [IO.Path]::DirectorySeparatorChar
    if (-not $cleanupPath.StartsWith($allowedRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture path escaped backups.' }
    try { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($cleanupPath,'OnlyErrorDialogs','SendToRecycleBin') }
    catch { Write-Warning "Test fixtures retained: $cleanupPath" }
}
