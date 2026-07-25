[CmdletBinding()]
param(
    [string]$DotmPath = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "ASSERTION FAILED: $Message"
    }
}

function Assert-Contains {
    param(
        [string]$Text,
        [string]$Needle,
        [string]$Message
    )

    Assert-True ($Text.Contains($Needle)) $Message
}

function Read-VbaSource {
    param([string]$Path)

    $vbaEncoding = [System.Text.Encoding]::GetEncoding(936)
    return $vbaEncoding.GetString([System.IO.File]::ReadAllBytes($Path))
}

Write-Output "Running QuickToolbar source checks"
$formPath = Join-Path $repoRoot "userforms\frmQuickToolbar.frm"
$frxPath = Join-Path $repoRoot "userforms\frmQuickToolbar.frx"
$modulePath = Join-Path $repoRoot "modules\Mod_QuickToolbar.bas"
$ribbonPath = Join-Path $repoRoot "dotm\customUI\customUI14.xml"

Assert-True (Test-Path -LiteralPath $formPath) "Quick toolbar form source should exist."
Assert-True (Test-Path -LiteralPath $frxPath) "Quick toolbar form binary resource should exist."
Assert-True (Test-Path -LiteralPath $modulePath) "Quick toolbar module should exist."
Assert-True ((Get-Item -LiteralPath $frxPath).Length -gt 15000) "Quick toolbar FRX should contain eight embedded icon pictures."

$formText = Read-VbaSource -Path $formPath
$moduleText = Read-VbaSource -Path $modulePath
$ribbonText = Get-Content -LiteralPath $ribbonPath -Raw

Assert-Contains $formText 'OleObjectBlob   =   "frmQuickToolbar.frx":0000' "Form should reference its .frx resource."
Assert-Contains $formText "ShowModal       =   0" "Quick toolbar should be modeless."
Assert-Contains $formText "ClientHeight    =   4700" "Quick toolbar should use a compact vertical client area."
Assert-Contains $formText "ClientWidth     =   1600" "Quick toolbar should use a compact narrow client area."
Assert-Contains $formText 'Me.Caption = "RATools"' "Quick toolbar should use a short title."
Assert-Contains $formText '.Left = (Me.InsideWidth - .Width) / 2' "Quick toolbar icons should be centered from the live client width."
Assert-Contains $formText '.Top = BUTTON_TOP + rowIndex * BUTTON_STEP' "Quick toolbar icons should be arranged in one vertical column."
Assert-Contains $formText 'Private Const BUTTON_WIDTH As Single = 28' "Icon tiles should use a compact width."
Assert-Contains $formText 'Private Const BUTTON_HEIGHT As Single = 25' "Icon tiles should use a compact height."
Assert-Contains $formText 'Private Const BUTTON_STEP As Single = 27' "Icon tiles should use a small vertical gap."
Assert-Contains $formText 'ByVal buttonItem As MSForms.Label' "Quick toolbar controls should use flat label tiles instead of raised command buttons."
Assert-Contains $formText '.BorderStyle = fmBorderStyleSingle' "Icon tiles should use a thin flat border."
Assert-Contains $formText '.PicturePosition = fmPicturePositionCenter' "Icons should be centered in their buttons."
Assert-Contains $formText '图标直接复用 Ribbon 资源' "Quick toolbar should document reuse of the Ribbon icon sources."
Assert-Contains $formText '0, "H1"' "Heading 1 should use the existing H1 Ribbon icon."
Assert-Contains $formText '6, "imageMso:TableAutoFitWindow"' "Table auto-fit should use the existing Ribbon imageMso icon."
Assert-Contains $formText '7, "idMso:Subscript"' "Terms should use the existing Ribbon subscript icon."
Assert-Contains $formText 'ConfigureButton Me.btnAutoFitTable, "表格适应"' "Auto-fit icon button should remain wired."
Assert-Contains $moduleText "Public Sub ToggleQuickToolbar" "Ribbon toggle callback should exist."
Assert-Contains $moduleText "Public Sub RunQuickToolbarAction" "Action dispatcher should exist."
Assert-Contains $moduleText 'Private Const QUICK_TOOLBAR_TITLE As String = "RATools 快捷工具栏"' "Quick toolbar messages should keep a readable Chinese title."
Assert-Contains $formText "Private Sub btnHeading1_Click()" "Heading 1 button click handler should exist."
Assert-Contains $formText "Private Sub btnNormalizeTerms_Click()" "Scientific term button click handler should exist."
Assert-Contains $ribbonText 'onAction="ToggleQuickToolbar"' "Ribbon should expose the quick toolbar callback."

$expectedActions = @(
    "style_heading1",
    "style_heading2",
    "style_heading3",
    "style_body",
    "text_blue",
    "page_break_before",
    "autofit_table",
    "normalize_terms"
)
foreach ($action in $expectedActions) {
    Assert-Contains $moduleText ('"' + $action + '"') "Missing quick toolbar action: $action"
}

Write-Output "Running QuickToolbar Word COM source-import smoke"
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1
    $doc = $word.Documents.Add()

    $sourceFiles = @(
        (Join-Path $repoRoot "modules\Mod_QuickToolbar.bas"),
        (Join-Path $repoRoot "userforms\frmQuickToolbar.frm")
    )

    foreach ($sourceFile in $sourceFiles) {
        try {
            [void]$doc.VBProject.VBComponents.Import($sourceFile)
        }
        catch {
            throw "Could not import ${sourceFile}: $($_.Exception.Message)"
        }
    }

    $stubComponent = $doc.VBProject.VBComponents.Add(1)
    $stubComponent.Name = "QuickToolbarTestStubs"
$stubComponent.CodeModule.AddFromString(@'
Option Explicit

Public LastQuickAction As String

Public Sub ApplyRAToolsStyle(ByVal uiTagName As String)
    LastQuickAction = "style:" & uiTagName
End Sub

Public Sub SetTextBlue(ByVal control As IRibbonControl)
    LastQuickAction = "text_blue"
End Sub

Public Sub TogglePageBreakBefore(ByVal control As IRibbonControl)
    LastQuickAction = "page_break_before"
End Sub

Public Sub AutoFitTableWindow(ByVal control As IRibbonControl)
    LastQuickAction = "autofit_table"
End Sub

Public Sub NormalizeScientificTerms()
    LastQuickAction = "normalize_terms"
End Sub

Public Function GetLastQuickAction() As String
    GetLastQuickAction = LastQuickAction
End Function

Public Function GetQuickToolbarCaptionForTest() As String
    Load frmQuickToolbar
    GetQuickToolbarCaptionForTest = frmQuickToolbar.Caption
End Function

Public Function GetQuickToolbarFormSizeForTest() As String
    Load frmQuickToolbar
    GetQuickToolbarFormSizeForTest = CStr(frmQuickToolbar.Width) & "|" & _
        CStr(frmQuickToolbar.Height) & "|" & CStr(frmQuickToolbar.InsideWidth) & "|" & _
        CStr(frmQuickToolbar.InsideHeight)
End Function

Public Function GetQuickToolbarControlStateForTest(ByVal controlName As String) As String
    Dim toolbarControl As Object
    Dim pictureItem As Object
    Dim hasPicture As Boolean

    Load frmQuickToolbar
    Set toolbarControl = frmQuickToolbar.Controls(controlName)
    On Error Resume Next
    Set pictureItem = toolbarControl.Picture
    hasPicture = Not (pictureItem Is Nothing)
    On Error GoTo 0
    GetQuickToolbarControlStateForTest = toolbarControl.Caption & "|" & _
        CStr(toolbarControl.Left) & "|" & CStr(toolbarControl.Top) & "|" & _
        CStr(toolbarControl.Width) & "|" & CStr(toolbarControl.Height) & "|" & _
        toolbarControl.ControlTipText & "|" & CStr(hasPicture) & "|" & _
        TypeName(toolbarControl) & "|" & toolbarControl.Tag
End Function
'@)

    $buttonCount = [int]$word.Run("GetQuickToolbarButtonCount")
    Assert-True ($buttonCount -eq 8) "Quick toolbar should contain eight design-time icon controls; got $buttonCount."

    $toolbarCaption = [string]$word.Run("GetQuickToolbarCaptionForTest")
    Assert-True ($toolbarCaption -eq "RATools") "Quick toolbar should render the short title; got $toolbarCaption."

    $formSize = ([string]$word.Run("GetQuickToolbarFormSizeForTest")).Split("|")
    Assert-True ($formSize.Count -eq 4) "Quick toolbar form and client sizes should be returned for verification."
    $formWidth = [double]$formSize[0]
    $formHeight = [double]$formSize[1]
    $insideWidth = [double]$formSize[2]
    $insideHeight = [double]$formSize[3]
    Assert-True ($formWidth -le 105) "Quick toolbar should remain compact even with the Windows title-bar minimum width; got $formWidth."
    Assert-True ($formHeight -le 260) "Quick toolbar should use a compact height; got $formHeight."
    Assert-True ($formWidth -lt $formHeight) "Quick toolbar should be narrower than it is tall."

    $expectedButtons = [ordered]@{
        btnHeading1      = "标题 1"
        btnHeading2      = "标题 2"
        btnHeading3      = "标题 3"
        btnBody          = "正文"
        btnTextBlue      = "设为蓝色"
        btnPageBreakBefore = "段前分页"
        btnAutoFitTable  = "表格适应"
        btnNormalizeTerms = "术语下标"
    }
    $expectedIconSources = [ordered]@{
        btnHeading1      = "H1"
        btnHeading2      = "H2"
        btnHeading3      = "H3"
        btnBody          = "BodyText"
        btnTextBlue      = "SetTextBlue"
        btnPageBreakBefore = "PageBreakBefore"
        btnAutoFitTable  = "imageMso:TableAutoFitWindow"
        btnNormalizeTerms = "idMso:Subscript"
    }

    $previousTop = -1.0
    $previousBottom = -1.0
    foreach ($buttonEntry in $expectedButtons.GetEnumerator()) {
        $controlName = [string]$buttonEntry.Key
        $state = ([string]$word.Run("GetQuickToolbarControlStateForTest", [ref]$controlName)).Split("|", 9)
        Assert-True ($state.Count -eq 9) "Could not inspect runtime state for $controlName."
        Assert-True ([string]::IsNullOrEmpty($state[0])) "$controlName should use an icon instead of visible button text; got $($state[0])"
        Assert-True ($state[6] -eq "True") "$controlName should expose a picture icon at runtime."
        Assert-True ($state[7] -eq "Label") "$controlName should render as a flat label tile; got $($state[7])."
        Assert-True ($state[8] -eq [string]$expectedIconSources[$controlName]) "$controlName should reuse the matching Ribbon icon source; got $($state[8])."

        $left = [double]$state[1]
        $top = [double]$state[2]
        $width = [double]$state[3]
        $height = [double]$state[4]
        $controlCenter = $left + ($width / 2.0)
        $clientCenter = $insideWidth / 2.0
        Assert-True ([Math]::Abs($controlCenter - $clientCenter) -lt 0.15) "$controlName should be exactly centered in the live client width."
        Assert-True ($width -le 30) "$controlName should use a compact width; got $width."
        Assert-True ($height -le 27) "$controlName should use a compact height; got $height."
        Assert-True ($top -gt $previousTop) "$controlName should appear below the previous button."
        if ($previousBottom -ge 0) {
            Assert-True (($top - $previousBottom) -le 3.0) "$controlName should use only a small vertical gap."
        }
        Assert-True (($top + $height) -le ($insideHeight + 0.5)) "$controlName should fit completely inside the compact form."
        Assert-True ($state[5].StartsWith([string]$buttonEntry.Value)) "$controlName should have a readable named tooltip."
        $previousTop = $top
        $previousBottom = $top + $height
    }

    [void]$word.Run("ReleaseQuickToolbarForTest")
    $toolbarComponent = $doc.VBProject.VBComponents.Item("frmQuickToolbar")
    $expectedDesignerLeft = $null
    $previousDesignerTop = -1.0
    foreach ($buttonEntry in $expectedButtons.GetEnumerator()) {
        $designerControl = $toolbarComponent.Designer.Controls.Item([string]$buttonEntry.Key)
        Assert-True ($designerControl.Caption -eq [string]$buttonEntry.Value) "Design-time caption should be readable for $($buttonEntry.Key)."
        Assert-True ($designerControl.Tag -eq [string]$expectedIconSources[[string]$buttonEntry.Key]) "Design-time icon source should match Ribbon mapping for $($buttonEntry.Key)."
        Assert-True ([double]$designerControl.Width -le 30) "$($buttonEntry.Key) should use a compact design-time width."
        Assert-True ([double]$designerControl.Height -le 27) "$($buttonEntry.Key) should use a compact design-time height."
        if ($null -eq $expectedDesignerLeft) {
            $expectedDesignerLeft = [double]$designerControl.Left
        }
        Assert-True ([Math]::Abs([double]$designerControl.Left - $expectedDesignerLeft) -lt 0.1) "$($buttonEntry.Key) should align in the design-time column."
        Assert-True ([double]$designerControl.Top -gt $previousDesignerTop) "$($buttonEntry.Key) should be vertically ordered at design time."
        $previousDesignerTop = [double]$designerControl.Top
    }

    $actionCases = @(
        @{ Key = "style_heading1"; Expected = "style:标题1-F" },
        @{ Key = "style_heading2"; Expected = "style:标题2-F" },
        @{ Key = "style_heading3"; Expected = "style:标题3-F" },
        @{ Key = "style_body"; Expected = "style:正文-F" },
        @{ Key = "text_blue"; Expected = "text_blue" },
        @{ Key = "page_break_before"; Expected = "page_break_before" },
        @{ Key = "autofit_table"; Expected = "autofit_table" },
        @{ Key = "normalize_terms"; Expected = "normalize_terms" }
    )
    foreach ($actionCase in $actionCases) {
        $actionKey = [string]$actionCase.Key
        [void]$word.Run("RunQuickToolbarAction", [ref]$actionKey)
        $actualAction = [string]$word.Run("GetLastQuickAction")
        Assert-True ($actualAction -eq $actionCase.Expected) "Unexpected dispatch result for $($actionCase.Key): $actualAction"
    }

    [void]$word.Run("ReleaseQuickToolbarForTest")

    if (-not [string]::IsNullOrWhiteSpace($DotmPath)) {
        Assert-True (Test-Path -LiteralPath $DotmPath -PathType Leaf) "Built dotm should exist: $DotmPath"
        $resolvedDotmPath = (Resolve-Path -LiteralPath $DotmPath).Path

        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $archive = [System.IO.Compression.ZipFile]::OpenRead($resolvedDotmPath)
        try {
            $ribbonEntry = $archive.GetEntry("customUI/customUI14.xml")
            Assert-True ($null -ne $ribbonEntry) "Built dotm should contain customUI/customUI14.xml."
            $reader = New-Object System.IO.StreamReader($ribbonEntry.Open(), [System.Text.Encoding]::UTF8)
            try {
                $artifactRibbonText = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }
            Assert-Contains $artifactRibbonText 'id="btnQuickToolbar"' "Built dotm Ribbon should contain btnQuickToolbar."
            Assert-Contains $artifactRibbonText 'onAction="ToggleQuickToolbar"' "Built dotm Ribbon should call ToggleQuickToolbar."
        }
        finally {
            $archive.Dispose()
        }

        $doc.Close($false) | Out-Null
        $doc = $null

        Write-Output "Inspecting built QuickToolbar artifact"
        $doc = $word.Documents.Open($resolvedDotmPath, $false, $true, $false)
        $componentNames = @($doc.VBProject.VBComponents | ForEach-Object { $_.Name })
        Assert-True ($componentNames -contains "Mod_QuickToolbar") "Built dotm should contain Mod_QuickToolbar."
        Assert-True ($componentNames -contains "frmQuickToolbar") "Built dotm should contain frmQuickToolbar."

        $toolbarComponent = $doc.VBProject.VBComponents.Item("frmQuickToolbar")
        $artifactButtonCount = [int]$toolbarComponent.Designer.Controls.Count
        Assert-True ($artifactButtonCount -eq 8) "Built dotm toolbar should contain eight buttons; got $artifactButtonCount."

        $artifactCode = $toolbarComponent.CodeModule.Lines(1, $toolbarComponent.CodeModule.CountOfLines)
        Assert-Contains $artifactCode 'Me.Caption = "RATools"' "Built dotm should preserve the compact toolbar title."
        Assert-Contains $artifactCode 'ConfigureButton Me.btnNormalizeTerms, "术语下标"' "Built dotm should preserve the terms button wiring."
        Assert-Contains $artifactCode '.Left = (Me.InsideWidth - .Width) / 2' "Built dotm should preserve exact dynamic centering."
        Assert-Contains $artifactCode 'ByVal buttonItem As MSForms.Label' "Built dotm should preserve flat icon tiles."
        Assert-Contains $artifactCode 'imageMso:TableAutoFitWindow' "Built dotm should preserve the Ribbon auto-fit icon mapping."
        Assert-Contains $artifactCode 'idMso:Subscript' "Built dotm should preserve the Ribbon subscript icon mapping."
        $expectedArtifactLeft = $null
        $previousArtifactTop = -1.0
        foreach ($buttonEntry in $expectedButtons.GetEnumerator()) {
            $artifactControl = $toolbarComponent.Designer.Controls.Item([string]$buttonEntry.Key)
            Assert-True ($artifactControl.Caption -eq [string]$buttonEntry.Value) "Built dotm caption should be readable for $($buttonEntry.Key)."
            Assert-True ($artifactControl.Tag -eq [string]$expectedIconSources[[string]$buttonEntry.Key]) "Built dotm icon source should match Ribbon mapping for $($buttonEntry.Key)."
            Assert-True ([double]$artifactControl.Width -le 30) "Built dotm should preserve the compact width for $($buttonEntry.Key)."
            Assert-True ([double]$artifactControl.Height -le 27) "Built dotm should preserve the compact height for $($buttonEntry.Key)."
            if ($null -eq $expectedArtifactLeft) {
                $expectedArtifactLeft = [double]$artifactControl.Left
            }
            Assert-True ([Math]::Abs([double]$artifactControl.Left - $expectedArtifactLeft) -lt 0.1) "Built dotm should align $($buttonEntry.Key) in one column."
            Assert-True ([double]$artifactControl.Top -gt $previousArtifactTop) "Built dotm should vertically order $($buttonEntry.Key)."
            $previousArtifactTop = [double]$artifactControl.Top
        }
    }

    Write-Output "PASS QuickToolbar.Tests"
}
catch {
    throw "Quick toolbar Word COM source-import smoke failed. Enable 'Trust access to the VBA project object model' if required. $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)"
}
finally {
    if ($doc -ne $null) {
        $doc.Close($false) | Out-Null
    }
    if ($word -ne $null) {
        $word.Quit() | Out-Null
    }
}
