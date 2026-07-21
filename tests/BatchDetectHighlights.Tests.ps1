$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$modulePath = Join-Path $repoRoot "modules\Mod_BatchDetectHighlights.bas"

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)

    if ($Expected -ne $Actual) {
        throw "$Message Expected <$Expected>, got <$Actual>."
    }
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1

    $doc = $word.Documents.Add()
    $component = $doc.VBProject.VBComponents.Import($modulePath)
    $component.CodeModule.AddFromString(@'
Public Sub RAToolsHighlightTestRunner()
    Dim hit As Range
    Dim detectionError As String

    On Error Resume Next
    ThisDocument.Variables("RAToolsHighlightResult").Delete
    ThisDocument.Variables("RAToolsHighlightText").Delete
    ThisDocument.Variables("RAToolsHighlightSummary").Delete
    ThisDocument.Variables("RAToolsHighlightError").Delete
    On Error GoTo 0

    If TryFindFirstHighlight(ThisDocument, hit, detectionError) Then
        ThisDocument.Variables.Add "RAToolsHighlightResult", "True"
        ThisDocument.Variables.Add "RAToolsHighlightText", hit.Text
        ThisDocument.Variables.Add "RAToolsHighlightSummary", GetHighlightSummary(hit)
    Else
        ThisDocument.Variables.Add "RAToolsHighlightResult", "False"
    End If
    ThisDocument.Variables.Add "RAToolsHighlightError", detectionError
End Sub

Public Sub RAToolsHighlightErrorTestRunner()
    Dim hit As Range
    Dim detectionError As String
    Dim missingDocument As Document

    On Error Resume Next
    ThisDocument.Variables("RAToolsHighlightErrorResult").Delete
    ThisDocument.Variables("RAToolsHighlightExpectedError").Delete
    On Error GoTo 0

    ThisDocument.Variables.Add "RAToolsHighlightErrorResult", _
        CStr(TryFindFirstHighlight(missingDocument, hit, detectionError))
    ThisDocument.Variables.Add "RAToolsHighlightExpectedError", detectionError
End Sub
'@)

    function Invoke-HighlightCheck {
        function Get-DocumentVariableValue {
            param([string]$Name)

            try {
                return $doc.Variables.Item($Name).Value
            }
            catch {
                return ""
            }
        }

        $null = $word.Run("RAToolsHighlightTestRunner")
        return [pscustomobject]@{
            Found = [bool]::Parse($doc.Variables.Item("RAToolsHighlightResult").Value)
            Text = Get-DocumentVariableValue "RAToolsHighlightText"
            Summary = Get-DocumentVariableValue "RAToolsHighlightSummary"
            Error = [string]$doc.Variables.Item("RAToolsHighlightError").Value
        }
    }

    $doc.Range().Text = "plain text"
    $doc.Range().Font.Hidden = 0
    $doc.Range().HighlightColorIndex = 0
    Assert-Equal $false (Invoke-HighlightCheck).Found "Plain text should not match."

    $doc.Range().Text = "plain text"
    $doc.Range().Font.Hidden = 0
    $doc.Range().HighlightColorIndex = 0
    $doc.Range($doc.Content.End - 1, $doc.Content.End).HighlightColorIndex = 7
    Assert-Equal $false (Invoke-HighlightCheck).Found "Highlighted paragraph mark should not match."

    $doc.Range().Text = "a`t b"
    $doc.Range().Font.Hidden = 0
    $doc.Range().HighlightColorIndex = 0
    $doc.Range(1, 3).HighlightColorIndex = 7
    Assert-Equal $false (Invoke-HighlightCheck).Found "Highlighted whitespace should not match."

    $doc.Range().Text = "shown hidden"
    $doc.Range().Font.Hidden = 0
    $doc.Range().HighlightColorIndex = 0
    $doc.Range(6, 12).Font.Hidden = -1
    $doc.Range(6, 12).HighlightColorIndex = 7
    Assert-Equal $false (Invoke-HighlightCheck).Found "Highlighted hidden text should not match."

    $doc.Range().Text = " visible"
    $doc.Range().Font.Hidden = 0
    $doc.Range().HighlightColorIndex = 0
    $doc.Range(0, 1).HighlightColorIndex = 7
    $doc.Range(1, 8).HighlightColorIndex = 7
    $result = Invoke-HighlightCheck
    Assert-Equal $true $result.Found "Visible highlight after an ignored match should be found."
    Assert-Equal "visible" $result.Text "The visible highlighted run should be returned."
    $expectedSummary = "$([char]0x6B63)$([char]0x6587)$([char]0xFF1A)visible"
    Assert-Equal $expectedSummary $result.Summary "The report summary should identify body text."
    Assert-Equal "" $result.Error "Successful detection should not report an error."

    $doc.Range().Text = "plain body"
    $doc.Range().Font.Hidden = 0
    $doc.Range().HighlightColorIndex = 0
    $footnote = $doc.Footnotes.Add($doc.Range(1, 1), "", "footnote hit")
    $footnote.Range.HighlightColorIndex = 7
    $result = Invoke-HighlightCheck
    Assert-Equal $true $result.Found "Visible footnote highlight should be found."
    Assert-Equal "footnote" $result.Text "The first visible footnote highlight run should be returned."
    $expectedSummary = "$([char]0x811A)$([char]0x6CE8)$([char]0xFF1A)footnote"
    Assert-Equal $expectedSummary $result.Summary "The report summary should identify a footnote."

    $null = $word.Run("RAToolsHighlightErrorTestRunner")
    $errorResult = [bool]::Parse($doc.Variables.Item("RAToolsHighlightErrorResult").Value)
    $expectedError = [string]$doc.Variables.Item("RAToolsHighlightExpectedError").Value
    Assert-Equal $false $errorResult "Detection errors must not return a stale true result."
    if ([string]::IsNullOrWhiteSpace($expectedError)) {
        throw "Detection errors should provide an error message."
    }

    Write-Host "PASS BatchDetectHighlights.Tests"
}
finally {
    if ($null -ne $doc) {
        $doc.Close($false) | Out-Null
    }
    if ($null -ne $word) {
        $word.Quit() | Out-Null
    }
}
