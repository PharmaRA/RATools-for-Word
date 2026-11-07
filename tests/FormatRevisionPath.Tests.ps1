$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = Join-Path (Join-Path $repo 'backups') ('RATools_PathRepair_' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory((Join-Path $root 'ProjectA'))
[void][IO.Directory]::CreateDirectory((Join-Path $root 'ProjectB'))
$files = @('Report.docx','ProjectA\Report.docx','ProjectB\Report.docx')
foreach ($name in $files) { [IO.File]::WriteAllText((Join-Path $root $name), 'unchanged fixture') }
$gbk = [Text.Encoding]::GetEncoding(936)
$macro = [IO.File]::ReadAllText((Join-Path $repo 'modules\Macro_AcceptFormatRevisions.bas'), $gbk)
$core = [IO.File]::ReadAllText((Join-Path $repo 'modules\Core_Files.bas'), $gbk)
function Replace-Procedure($text, $name, $replacement) {
    $pattern = '(?ms)^(?:Private|Public)?\s*(?:Sub|Function) ' + $name + '\(.*?^End (?:Sub|Function)'
    return [regex]::Replace($text, $pattern, [Text.RegularExpressions.MatchEvaluator]{ param($m) $replacement })
}
# Substitute only external mapping discovery, the active URL and irreversible work.
# The production path selection and exact mapping algorithms execute unchanged.
$macro = $macro.Replace('docPath = doc.FullName', 'docPath = repairUrl')
$macro = $macro.Replace('Option Explicit', "Option Explicit`r`nPrivate repairUrl As String`r`nPrivate repairResult As String")
$macro = Replace-Procedure $macro 'AcceptFormatRevisionsByOoxml' @'
Private Sub AcceptFormatRevisionsByOoxml(ByVal docPath As String)
    repairResult = "FILE:" & docPath
End Sub
'@
$macro = Replace-Procedure $macro 'AcceptFormatRevisionsByWordApi' @'
Private Sub AcceptFormatRevisionsByWordApi(ByVal doc As Document)
    If doc Is ActiveDocument Then repairResult = "WORD_API"
End Sub
'@
$macro = Replace-Procedure $macro 'ListOneDriveRoots' @"
Private Function ListOneDriveRoots(ByVal fso As Object) As Collection
    Dim roots As New Collection
    roots.Add "$root"
    Set ListOneDriveRoots = roots
End Function
"@
$core = Replace-Procedure $core 'LoadOneDriveSyncMappings' @"
Private Sub LoadOneDriveSyncMappings(ByRef urls As Collection, ByRef roots As Collection)
    urls.Add "https://mapped.invalid/Documents"
    roots.Add "$root"
End Sub
"@
$macro = [regex]::Replace($macro, '\bMsgBox\b', 'CapturePathMessage')
$macro += @'

Private Function CapturePathMessage(ByVal text As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
End Function
Public Function MissingMappingProbe() As String
    repairUrl = "https://unmapped.invalid/ProjectA/Report.docx"
    repairResult = ""
    Core_AcceptFormatRevisions
    MissingMappingProbe = repairResult
End Function
Public Function ExactMappingProbe() As String
    repairUrl = "https://mapped.invalid/Documents/ProjectA/Report.docx"
    repairResult = ""
    Core_AcceptFormatRevisions
    ExactMappingProbe = repairResult
End Function
Public Function MissingFileProbe() As String
    repairUrl = "https://mapped.invalid/Documents/Missing/Report.docx"
    repairResult = ""
    Core_AcceptFormatRevisions
    MissingFileProbe = repairResult
End Function
'@
$word = $null; $doc = $null
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
Add-Type -AssemblyName Microsoft.VisualBasic
try {
    [IO.File]::WriteAllText((Join-Path $root 'Macro_AcceptFormatRevisions.bas'), $macro, $gbk)
    [IO.File]::WriteAllText((Join-Path $root 'Core_Files.bas'), $core, $gbk)
    $word = Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $doc = $word.Documents.Add()
    [void]$doc.VBProject.VBComponents.Import((Join-Path $root 'Core_Files.bas'))
    [void]$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Core_Session.bas'))
    [void]$doc.VBProject.VBComponents.Import((Join-Path $root 'Macro_AcceptFormatRevisions.bas'))
    $result = $word.Run('MissingMappingProbe')
    if ($result -ne 'WORD_API') { throw "Unmapped URL must use current document, not a same-name file. Actual: $result" }
    Write-Host 'PASS missing mapping uses current document'
    $result = $word.Run('ExactMappingProbe')
    if ($result -ne ('FILE:' + (Join-Path $root 'ProjectA\Report.docx'))) { throw 'Exact mapping must preserve project identity' }
    Write-Host 'PASS exact mapping uses full relative path'
    if ($word.Run('MissingFileProbe') -ne 'WORD_API') { throw 'Missing exact file must not choose a different same-name file' }
    Write-Host 'PASS missing exact file uses current document'
    foreach ($name in $files) {
        if ([IO.File]::ReadAllText((Join-Path $root $name)) -ne 'unchanged fixture') { throw 'An unrelated fixture was changed' }
    }
} finally {
    if ($null -ne $doc) { $doc.Close(0) }
    Stop-RAToolsWordSession -Word $word
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($root, 'OnlyErrorDialogs', 'SendToRecycleBin')
}
Write-Host 'PASS FormatRevisionPath.Tests'
