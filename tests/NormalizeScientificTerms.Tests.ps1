$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$modulePath = Join-Path $repoRoot "modules\Mod_NormalizeScientificTerms.bas"
$logPath = Join-Path ([System.IO.Path]::GetTempPath()) "RATools_NormalizeScientificTerms_Test.log"

function Write-Step {
    param([string]$Message)

    $line = "{0} {1}" -f (Get-Date -Format "HH:mm:ss"), $Message
    Add-Content -LiteralPath $logPath -Value $line
    Write-Host $line
}

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Equal {
    param(
        $Expected,
        $Actual,
        [string]$Message
    )

    if ($Expected -ne $Actual) {
        throw "$Message Expected <$Expected>, got <$Actual>."
    }
}

function Assert-Subscript {
    param(
        $Document,
        [int]$Start,
        [int]$End,
        [bool]$Expected,
        [string]$Message
    )

    $range = $Document.Range($Start, $End)
    $actual = [int]$range.Font.Subscript
    $expectedValue = if ($Expected) { -1 } else { 0 }
    Assert-Equal $expectedValue $actual $Message
}

function Assert-TermFormatting {
    param(
        $Document,
        [string]$Text,
        [string]$Term,
        [int]$BaseLength
    )

    $start = $Text.IndexOf($Term)
    Assert-True ($start -ge 0) "Test setup could not find term <$Term>."

    Assert-Subscript $Document $start ($start + $BaseLength) $false "$Term base should not be subscript."
    Assert-Subscript $Document ($start + $BaseLength) ($start + $Term.Length) $true "$Term suffix should be subscript."
}

if (-not (Test-Path $modulePath)) {
    throw "Expected VBA module not found: $modulePath"
}

Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
Write-Step "Starting Word COM test"

$word = $null
$doc = $null

try {
    Write-Step "Creating Word.Application"
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1

    Write-Step "Creating document"
    $doc = $word.Documents.Add()

    try {
        Write-Step "Importing VBA module"
        [void]$doc.VBProject.VBComponents.Import($modulePath)
    }
    catch {
        throw "Could not import VBA module. Enable 'Trust access to the VBA project object model' in Word Trust Center to run this test. $($_.Exception.Message)"
    }

    $infinity = [char]0x221E
    $tau = [char]0x03C4
    $aucInfinity = "AUC0-$infinity"
    $aucTau = "AUC0-$tau"
    $text = "STD10 Cmax Cmin Tmax Tmin AUC0-t AUC0-inf AUCtau AUClast t1/2 " +
        "AUC0-last AUC0-24h $aucInfinity $aucTau AUC0-tau " +
        "C0 Ct Css Cavg Cave Ctrough Vd Vz/F Vd/F CLz/F CL/F MYSTD20 Cmaxed"
    Write-Step "Writing full-document test text"
    $doc.Range().Text = $text

    Write-Step "Running NormalizeScientificTermsInRange"
    $fullRange = $doc.Range()
    $count = $word.Run("NormalizeScientificTermsInRange", [ref]$fullRange)
    Assert-Equal 25 $count "Unexpected formatted term count."

    Assert-TermFormatting $doc $text "STD10" 3
    Assert-TermFormatting $doc $text "Cmax" 1
    Assert-TermFormatting $doc $text "Cmin" 1
    Assert-TermFormatting $doc $text "Tmax" 1
    Assert-TermFormatting $doc $text "Tmin" 1
    Assert-TermFormatting $doc $text "AUC0-t" 3
    Assert-TermFormatting $doc $text "AUC0-inf" 3
    Assert-TermFormatting $doc $text "AUCtau" 3
    Assert-TermFormatting $doc $text "AUClast" 3
    Assert-TermFormatting $doc $text "t1/2" 1
    Assert-TermFormatting $doc $text "AUC0-last" 3
    Assert-TermFormatting $doc $text "AUC0-24h" 3
    Assert-TermFormatting $doc $text $aucInfinity 3
    Assert-TermFormatting $doc $text $aucTau 3
    Assert-TermFormatting $doc $text "AUC0-tau" 3
    Assert-TermFormatting $doc $text "C0" 1
    Assert-TermFormatting $doc $text "Ct" 1
    Assert-TermFormatting $doc $text "Css" 1
    Assert-TermFormatting $doc $text "Cavg" 1
    Assert-TermFormatting $doc $text "Cave" 1
    Assert-TermFormatting $doc $text "Ctrough" 1
    Assert-TermFormatting $doc $text "Vd" 1

    $vzStart = $text.IndexOf("Vz/F")
    Assert-Subscript $doc $vzStart ($vzStart + 1) $false "Vz/F V should not be subscript."
    Assert-Subscript $doc ($vzStart + 1) ($vzStart + 2) $true "Vz/F z should be subscript."
    Assert-Subscript $doc ($vzStart + 2) ($vzStart + 4) $false "Vz/F /F should not be subscript."

    $vdfStart = $text.IndexOf("Vd/F")
    Assert-Subscript $doc $vdfStart ($vdfStart + 1) $false "Vd/F V should not be subscript."
    Assert-Subscript $doc ($vdfStart + 1) ($vdfStart + 2) $true "Vd/F d should be subscript."
    Assert-Subscript $doc ($vdfStart + 2) ($vdfStart + 4) $false "Vd/F /F should not be subscript."

    $clzStart = $text.IndexOf("CLz/F")
    Assert-Subscript $doc $clzStart ($clzStart + 2) $false "CLz/F CL should not be subscript."
    Assert-Subscript $doc ($clzStart + 2) ($clzStart + 3) $true "CLz/F z should be subscript."
    Assert-Subscript $doc ($clzStart + 3) ($clzStart + 5) $false "CLz/F /F should not be subscript."

    $clStart = $text.IndexOf("CL/F")
    Assert-Subscript $doc $clStart ($clStart + 4) $false "CL/F should remain unchanged because it has no subscript part."

    $nonMatchStart = $text.IndexOf("MYSTD20")
    Assert-Subscript $doc ($nonMatchStart + 5) ($nonMatchStart + 7) $false "MYSTD20 should not be formatted inside a longer word."

    $cmaxedStart = $text.IndexOf("Cmaxed")
    Assert-Subscript $doc ($cmaxedStart + 1) ($cmaxedStart + 4) $false "Cmaxed should not format the max substring."

    $selectionText = "STD10 Cmax"
    Write-Step "Writing selection test text"
    $doc.Range().Text = $selectionText
    $doc.Range().Font.Subscript = $false

    $selectionStart = $selectionText.IndexOf("Cmax")
    $doc.Range($selectionStart, $selectionStart + 4).Select()

    Write-Step "Running NormalizeScientificTermsForCurrentTarget"
    $showMessage = $false
    $selectionCount = $word.Run("NormalizeScientificTermsForCurrentTarget", [ref]$showMessage)
    Assert-Equal 1 $selectionCount "Selection-only run should format exactly one term."

    Assert-Subscript $doc 3 5 $false "STD10 outside the selection should remain unchanged."
    Assert-TermFormatting $doc $selectionText "Cmax" 1

    Write-Step "PASS NormalizeScientificTerms"
}
finally {
    if ($doc -ne $null) {
        Write-Step "Closing document"
        $doc.Close($false) | Out-Null
    }
    if ($word -ne $null) {
        Write-Step "Quitting Word"
        $word.Quit() | Out-Null
    }
}
