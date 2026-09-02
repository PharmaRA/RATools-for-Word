[CmdletBinding()]
param(
    [string]$DotmPath = ""
)

$ErrorActionPreference = "Stop"

Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force
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

# 按钮定义单一事实源：scripts/QuickToolbarButtons.psd1
# 本测试的全部期望值（名称/提示/图标/布局/分发结果）由该文件推导，
# 修改按钮后无需同步改测试。
$buttonData = Import-RAToolsDataFile (Join-Path $repoRoot "scripts\QuickToolbarButtons.psd1")
$layoutSpec = $buttonData.Layout

function Get-ExpectedQuickToolbarLayout {
    param([int]$Index)

    $b = $buttonData.Buttons[$Index]
    return [pscustomobject]@{
        Row = [int]$b.Row
        Column = [int]$b.Column
        ColumnCount = [int]$b.ColumnCount
    }
}

$expectedItems = @(foreach ($b in $buttonData.Buttons) {
    $tag = switch ($b.Kind) {
        "custom"    { $b.Icon }
        "mso"       { "idMso:" + $b.Icon }
        "mso-image" { "imageMso:" + $b.Icon }
        default     { throw "Unknown button kind: $($b.Kind)" }
    }
    [pscustomobject]@{
        Name = $b.Name
        Caption = $b.Caption
        Tip = $b.Tip
        Tag = $tag
        Action = $b.Action
    }
})

$expectedDispatch = [ordered]@{}
foreach ($b in $buttonData.Buttons) {
    $expectedDispatch[$b.Action] = $b.Dispatch
}

# 由 psd1 布局参数推导按钮期望坐标（与窗体 GetButtonLeft/GetButtonTop 同公式）
function Get-ExpectedButtonGeometry {
    param([pscustomobject]$Layout, [double]$ClientWidth)

    $w = [double]$layoutSpec.ButtonWidth
    $gap = [double]$layoutSpec.ColumnGap
    $rowWidth = ($Layout.ColumnCount * $w) + (($Layout.ColumnCount - 1) * $gap)
    $left = (($ClientWidth - $rowWidth) / 2) + ($Layout.Column * ($w + $gap))
    $top = [double]$layoutSpec.ButtonTop + ($Layout.Row * [double]$layoutSpec.RowStep)
    foreach ($gapRow in $layoutSpec.GroupGapRows) {
        if ($Layout.Row -ge $gapRow) { $top += [double]$layoutSpec.GroupGap }
    }
    return [pscustomobject]@{ Left = $left; Top = $top }
}

Write-Output "Running QuickToolbar source checks"
$formPath = Join-Path $repoRoot "userforms\frmQuickToolbar.frm"
$frxPath = Join-Path $repoRoot "userforms\frmQuickToolbar.frx"
$modulePath = Join-Path $repoRoot "modules\UI_QuickToolbar.bas"
$actionModulePath = Join-Path $repoRoot "modules\UI_QuickToolbarActions.bas"
$styleNamesModulePath = Join-Path $repoRoot "modules\Engine_StyleNames.bas"
$windowModulePath = Join-Path $repoRoot "modules\Core_Window.bas"
$syncScriptPath = Join-Path $repoRoot "scripts\Sync-QuickToolbarForm.ps1"
$iconScriptPath = Join-Path $repoRoot "scripts\Generate-QuickToolbarIcon.ps1"
$ribbonPath = Join-Path $repoRoot "dotm\customUI\customUI14.xml"
$ribbonRelationshipsPath = Join-Path $repoRoot "dotm\customUI\_rels\customUI14.xml.rels"
$quickToolbarIconPath = Join-Path $repoRoot "dotm\customUI\images\QuickToolbar.png"

foreach ($requiredPath in @(
    $formPath,
    $frxPath,
    $modulePath,
    $actionModulePath,
    $styleNamesModulePath,
    $windowModulePath,
    $syncScriptPath,
    $iconScriptPath,
    $ribbonPath,
    $ribbonRelationshipsPath,
    $quickToolbarIconPath
)) {
    Assert-True (Test-Path -LiteralPath $requiredPath -PathType Leaf) "Required quick toolbar file should exist: $requiredPath"
}
Assert-True ((Get-Item -LiteralPath $frxPath).Length -gt 25000) "Quick toolbar FRX should contain all $($expectedItems.Count) embedded icon pictures."

$formText = Read-VbaSource -Path $formPath
$moduleText = Read-VbaSource -Path $modulePath
$actionModuleText = Read-VbaSource -Path $actionModulePath
$syncScriptText = Get-Content -LiteralPath $syncScriptPath -Raw -Encoding UTF8
$ribbonText = Get-Content -LiteralPath $ribbonPath -Raw
$ribbonRelationshipsText = Get-Content -LiteralPath $ribbonRelationshipsPath -Raw

Add-Type -AssemblyName System.Drawing
$quickToolbarIcon = [Drawing.Bitmap]::FromFile($quickToolbarIconPath)
try {
    Assert-True ($quickToolbarIcon.Width -eq 48) "Quick toolbar Ribbon icon should be 48px wide."
    Assert-True ($quickToolbarIcon.Height -eq 48) "Quick toolbar Ribbon icon should be 48px high."
    Assert-True ($quickToolbarIcon.PixelFormat.ToString().Contains("Argb")) "Quick toolbar Ribbon icon should preserve transparency."
}
finally {
    $quickToolbarIcon.Dispose()
}

Assert-Contains $formText 'OleObjectBlob   =   "frmQuickToolbar.frx":0000' "Form should reference its FRX resource."
Assert-Contains $formText "ShowModal       =   0" "Quick toolbar should be modeless."
Assert-Contains $formText 'Me.Caption = "RATools"' "Quick toolbar should use a short readable title."
Assert-Contains $formText "Private Const TOOLBAR_WIDTH As Single = $($layoutSpec.ToolbarWidth)" "Toolbar width should match QuickToolbarButtons.psd1."
Assert-Contains $formText "Private Const TOOLBAR_HEIGHT As Single = $($layoutSpec.ToolbarHeight)" "Toolbar height should match QuickToolbarButtons.psd1."
Assert-Contains $formText "Private Const BUTTON_WIDTH As Single = $($layoutSpec.ButtonWidth)" "Button width should match QuickToolbarButtons.psd1."
Assert-Contains $formText "Private Const BUTTON_HEIGHT As Single = $($layoutSpec.ButtonHeight)" "Button height should match QuickToolbarButtons.psd1."
Assert-Contains $formText "Private Const BUTTON_ROW_STEP As Single = $($layoutSpec.RowStep)" "Row step should match QuickToolbarButtons.psd1."
Assert-Contains $formText "Private Const BUTTON_COLUMN_GAP As Single = $($layoutSpec.ColumnGap)" "Column gap should match QuickToolbarButtons.psd1."
Assert-Contains $formText "If rowIndex >= 2 Then" "Toolbar should separate the heading and command groups."
Assert-Contains $formText "If rowIndex >= 4 Then" "Toolbar should separate the command and macro groups."
Assert-Contains $formText "GetButtonLeft = (Me.InsideWidth - rowWidth) / 2" "Each button row should be centered from the live client width."
Assert-Contains $formText "ByVal buttonItem As MSForms.Label" "Toolbar controls should use flat label tiles."
Assert-Contains $formText ".PicturePosition = fmPicturePositionCenter" "Toolbar icons should be centered inside each tile."
Assert-Contains $formText "Private Sub UserForm_MouseMove" "The form should hook mouse-move to raise the tooltip."
Assert-Contains $formText "_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" "Button mouse-move handlers should match the MSForms signature."
Assert-Contains $formText "RefreshTooltipWatch" "Form mouse-move handlers should call the refresh entry."
Assert-Contains $moduleText "Public Sub ToggleQuickToolbar" "Ribbon toggle callback should exist."
Assert-Contains $moduleText "Public Sub RunQuickToolbarAction" "Action dispatcher should exist."
foreach ($b in $buttonData.Buttons) {
    Assert-Contains $moduleText ('"' + $b.Action + '"') "Unified dispatcher should handle $($b.Action)."
}
Assert-Contains $actionModuleText "Public Sub ExecuteQuickToolbarMso" "Mso executor should exist."
Assert-Contains $actionModuleText 'Application.CommandBars.ExecuteMso commandId' "Built-in Word commands should use ExecuteMso."
Assert-Contains $ribbonText 'onAction="ToggleQuickToolbar"' "Ribbon should expose the quick toolbar callback."
$quickToolbarRibbonButton = [regex]::Match($ribbonText, '(?s)<button\s+id="btnQuickToolbar".*?/>').Value
$templateRibbonButton = [regex]::Match($ribbonText, '(?s)<button\s+id="btnLogo".*?/>').Value
Assert-True (-not [string]::IsNullOrWhiteSpace($quickToolbarRibbonButton)) "Ribbon should define btnQuickToolbar."
Assert-Contains $quickToolbarRibbonButton 'image="QuickToolbar"' "Quick toolbar Ribbon button should use its purpose-built icon."
Assert-Contains $templateRibbonButton 'image="Logo"' "Template loader should continue using the RATools Logo."
Assert-Contains $ribbonRelationshipsText 'Id="QuickToolbar"' "Ribbon relationships should register the quick toolbar icon."
Assert-Contains $ribbonRelationshipsText 'Target="images/QuickToolbar.png"' "Ribbon relationship should target QuickToolbar.png."

for ($i = 0; $i -lt $expectedItems.Count; $i++) {
    $item = $expectedItems[$i]
    $layout = Get-ExpectedQuickToolbarLayout -Index $i
    $expectedConfigure = "ConfigureButton Me.$($item.Name), $($layout.Row), $($layout.Column), $($layout.ColumnCount), $i"
    Assert-Contains $formText $expectedConfigure "Missing grid layout wiring for $($item.Name)."
    Assert-Contains $formText ("Private Sub " + $item.Name + "_Click()") "Missing click handler for $($item.Name)."
    Assert-Contains $formText ('RunQuickToolbarAction "' + $item.Action + '"') "Wrong action wiring for $($item.Name)."
}

Write-Output "Running QuickToolbar Word COM source-import smoke"
$word = $null
$doc = $null
try {
    $word = Start-RAToolsWordSession

    $doc = $word.Documents.Add()

    foreach ($sourceFile in @($modulePath, $actionModulePath, $styleNamesModulePath, $windowModulePath, $formPath)) {
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

Public Sub SetHyperlinksAndFieldsToBlue()
    LastQuickAction = "hyperlinks_fields_blue"
End Sub

Public Sub BatchAcceptAndClean()
    LastQuickAction = "accept_revisions_comments"
End Sub

Public Sub BatchDetectHighlights()
    LastQuickAction = "detect_highlights"
End Sub

Public Sub ResetQuickToolbarActionForTest()
    LastQuickAction = vbNullString
    SetQuickToolbarMsoTestMode True
End Sub

Public Function GetQuickToolbarActionForTest() As String
    If Len(LastQuickAction) > 0 Then
        GetQuickToolbarActionForTest = LastQuickAction
    ElseIf Len(GetLastQuickToolbarMsoForTest()) > 0 Then
        GetQuickToolbarActionForTest = "mso:" & GetLastQuickToolbarMsoForTest()
    End If
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
    Assert-True ($buttonCount -eq $expectedItems.Count) "Quick toolbar should contain $($expectedItems.Count) controls; got $buttonCount."

    $toolbarCaption = [string]$word.Run("GetQuickToolbarCaptionForTest")
    Assert-True ($toolbarCaption -eq "RATools") "Quick toolbar title should be readable; got $toolbarCaption."

    $formSize = ([string]$word.Run("GetQuickToolbarFormSizeForTest")).Split("|")
    Assert-True ($formSize.Count -eq 4) "Could not inspect quick toolbar runtime dimensions."
    $formWidth = [double]$formSize[0]
    $formHeight = [double]$formSize[1]
    $insideWidth = [double]$formSize[2]
    $insideHeight = [double]$formSize[3]
    Assert-True ([Math]::Abs($formWidth - $layoutSpec.ToolbarWidth) -lt 0.5) "Toolbar runtime width should be $($layoutSpec.ToolbarWidth)pt; got $formWidth."
    Assert-True ([Math]::Abs($formHeight - $layoutSpec.ToolbarHeight) -lt 0.5) "Toolbar runtime height should be $($layoutSpec.ToolbarHeight)pt; got $formHeight."
    Assert-True ($formHeight -lt 200) "Toolbar should remain a compact five-row grid."

    for ($i = 0; $i -lt $expectedItems.Count; $i++) {
        $item = $expectedItems[$i]
        $layout = Get-ExpectedQuickToolbarLayout -Index $i
        $controlName = [string]$item.Name
        $state = ([string]$word.Run("GetQuickToolbarControlStateForTest", [ref]$controlName)).Split("|", 9)
        Assert-True ($state.Count -eq 9) "Could not inspect runtime state for $controlName."
        Assert-True ([string]::IsNullOrEmpty($state[0])) "$controlName should show its embedded icon instead of text."
        Assert-True ($state[5] -eq ($item.Caption + "：" + $item.Tip)) "$controlName should keep its full Chinese tooltip."
        Assert-True ($state[6] -eq "True") "$controlName should expose an embedded picture."
        Assert-True ($state[7] -eq "Label") "$controlName should render as a flat Label tile."
        Assert-True ($state[8] -eq $item.Tag) "$controlName should reuse $($item.Tag); got $($state[8])."

        $left = [double]$state[1]
        $top = [double]$state[2]
        $width = [double]$state[3]
        $height = [double]$state[4]
        Assert-True ([Math]::Abs($width - $layoutSpec.ButtonWidth) -lt 0.1) "$controlName should be $($layoutSpec.ButtonWidth)pt wide; got $width."
        Assert-True ([Math]::Abs($height - $layoutSpec.ButtonHeight) -lt 0.1) "$controlName should be $($layoutSpec.ButtonHeight)pt high; got $height."

        $geometry = Get-ExpectedButtonGeometry -Layout $layout -ClientWidth $insideWidth
        $expectedLeft = $geometry.Left
        $expectedTop = $geometry.Top
        Assert-True ([Math]::Abs($left - $expectedLeft) -lt 0.15) "$controlName has an unexpected horizontal grid position: $left."
        Assert-True ([Math]::Abs($top - $expectedTop) -lt 0.15) "$controlName has an unexpected row position: $top."
        Assert-True ($left -ge -0.15) "$controlName should fit inside the left edge of the form."
        Assert-True (($left + $width) -le ($insideWidth + 0.15)) "$controlName should fit inside the right edge of the form."
        Assert-True (($top + $height) -le ($insideHeight + 0.15)) "$controlName should fit fully inside the form; bottom=$($top + $height), insideHeight=$insideHeight."
    }

    [void]$word.Run("ReleaseQuickToolbarForTest")
    $toolbarComponent = $doc.VBProject.VBComponents.Item("frmQuickToolbar")
    Assert-True ([int]$toolbarComponent.Designer.Controls.Count -eq $expectedItems.Count) "Designer should contain exactly $($expectedItems.Count) controls."
    for ($i = 0; $i -lt $expectedItems.Count; $i++) {
        $item = $expectedItems[$i]
        $control = $toolbarComponent.Designer.Controls.Item([string]$item.Name)
        $layout = Get-ExpectedQuickToolbarLayout -Index $i
        $geometry = Get-ExpectedButtonGeometry -Layout $layout -ClientWidth $layoutSpec.ToolbarWidth
        $expectedDesignerLeft = $geometry.Left
        $expectedDesignerTop = $geometry.Top
        Assert-True ($control.Caption -eq $item.Caption) "Design-time caption should be readable for $($item.Name)."
        Assert-True ($control.ControlTipText -eq ($item.Caption + "：" + $item.Tip)) "Design-time tooltip should be readable for $($item.Name)."
        Assert-True ($control.Tag -eq $item.Tag) "Design-time icon mapping should match for $($item.Name)."
        Assert-True ([Math]::Abs([double]$control.Left - $expectedDesignerLeft) -lt 0.15) "Design-time column should match for $($item.Name)."
        Assert-True ([Math]::Abs([double]$control.Top - $expectedDesignerTop) -lt 0.15) "Design-time row should match for $($item.Name)."
    }

    foreach ($actionCase in $expectedDispatch.GetEnumerator()) {
        [void]$word.Run("ResetQuickToolbarActionForTest")
        $actionKey = [string]$actionCase.Key
        [void]$word.Run("RunQuickToolbarAction", [ref]$actionKey)
        $actualAction = [string]$word.Run("GetQuickToolbarActionForTest")
        Assert-True ($actualAction -eq [string]$actionCase.Value) "Unexpected dispatch for ${actionKey}: $actualAction"
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
            $artifactQuickToolbarButton = [regex]::Match($artifactRibbonText, '(?s)<button\s+id="btnQuickToolbar".*?/>').Value
            Assert-Contains $artifactQuickToolbarButton 'image="QuickToolbar"' "Built quick toolbar button should use its purpose-built icon."

            $iconEntry = $archive.GetEntry("customUI/images/QuickToolbar.png")
            Assert-True ($null -ne $iconEntry) "Built dotm should contain QuickToolbar.png."
            Assert-True ($iconEntry.Length -gt 500) "Built QuickToolbar.png should contain rendered icon data."

            $relationshipsEntry = $archive.GetEntry("customUI/_rels/customUI14.xml.rels")
            Assert-True ($null -ne $relationshipsEntry) "Built dotm should contain customUI relationships."
            $reader = New-Object System.IO.StreamReader($relationshipsEntry.Open(), [System.Text.Encoding]::UTF8)
            try {
                $artifactRelationshipsText = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }
            Assert-Contains $artifactRelationshipsText 'Id="QuickToolbar"' "Built relationships should register QuickToolbar."
            Assert-Contains $artifactRelationshipsText 'Target="images/QuickToolbar.png"' "Built relationship should target QuickToolbar.png."
        }
        finally {
            $archive.Dispose()
        }

        $doc.Close($false) | Out-Null
        $doc = $null

        Write-Output "Inspecting built QuickToolbar artifact"
        $doc = $word.Documents.Open($resolvedDotmPath, $false, $true, $false)
        $componentNames = @($doc.VBProject.VBComponents | ForEach-Object { $_.Name })
        Assert-True ($componentNames -contains "UI_QuickToolbar") "Built dotm should contain UI_QuickToolbar."
        Assert-True ($componentNames -contains "UI_QuickToolbarActions") "Built dotm should contain UI_QuickToolbarActions."
        Assert-True ($componentNames -contains "frmQuickToolbar") "Built dotm should contain frmQuickToolbar."

        $artifactToolbar = $doc.VBProject.VBComponents.Item("frmQuickToolbar")
        Assert-True ([int]$artifactToolbar.Designer.Controls.Count -eq $expectedItems.Count) "Built toolbar should contain exactly $($expectedItems.Count) controls."
        $artifactCode = $artifactToolbar.CodeModule.Lines(1, $artifactToolbar.CodeModule.CountOfLines)
        Assert-Contains $artifactCode "GetButtonLeft = (Me.InsideWidth - rowWidth) / 2" "Built toolbar should preserve row centering."
        Assert-Contains $artifactCode "If rowIndex >= 2 Then" "Built toolbar should preserve the first group gap."
        Assert-Contains $artifactCode "If rowIndex >= 4 Then" "Built toolbar should preserve the second group gap."

        for ($i = 0; $i -lt $expectedItems.Count; $i++) {
            $item = $expectedItems[$i]
            $control = $artifactToolbar.Designer.Controls.Item([string]$item.Name)
            $layout = Get-ExpectedQuickToolbarLayout -Index $i
            $geometry = Get-ExpectedButtonGeometry -Layout $layout -ClientWidth $layoutSpec.ToolbarWidth
            $expectedArtifactLeft = $geometry.Left
            $expectedArtifactTop = $geometry.Top
            Assert-True ($control.Caption -eq $item.Caption) "Built caption should be readable for $($item.Name)."
            Assert-True ($control.ControlTipText -eq ($item.Caption + "：" + $item.Tip)) "Built tooltip should be readable for $($item.Name)."
            Assert-True ($control.Tag -eq $item.Tag) "Built icon mapping should match for $($item.Name)."
            Assert-True ([Math]::Abs([double]$control.Left - $expectedArtifactLeft) -lt 0.15) "Built column should match for $($item.Name)."
            Assert-True ([Math]::Abs([double]$control.Top - $expectedArtifactTop) -lt 0.15) "Built row should match for $($item.Name)."
        }

        $artifactTestComponent = $doc.VBProject.VBComponents.Add(1)
        $artifactTestComponent.Name = "QuickToolbarArtifactTest"
        $artifactTestComponent.CodeModule.AddFromString(@'
Option Explicit

Public Function GetBuiltQuickToolbarPictureState(ByVal controlName As String) As String
    Dim toolbarControl As Object
    Dim pictureItem As Object
    Dim hasPicture As Boolean

    Load frmQuickToolbar
    Set toolbarControl = frmQuickToolbar.Controls(controlName)
    On Error Resume Next
    Set pictureItem = toolbarControl.Picture
    hasPicture = Not (pictureItem Is Nothing)
    On Error GoTo 0
    GetBuiltQuickToolbarPictureState = toolbarControl.Caption & "|" & CStr(hasPicture)
End Function
'@)

        foreach ($item in $expectedItems) {
            $controlName = [string]$item.Name
            $pictureState = ([string]$word.Run("GetBuiltQuickToolbarPictureState", [ref]$controlName)).Split("|", 2)
            Assert-True ($pictureState.Count -eq 2) "Could not inspect built icon for $controlName."
            Assert-True ([string]::IsNullOrEmpty($pictureState[0])) "Built $controlName should show its icon instead of text."
            Assert-True ($pictureState[1] -eq "True") "Built $controlName should contain an embedded icon."
        }
        [void]$word.Run("ReleaseQuickToolbarForTest")

        $artifactActionModule = $doc.VBProject.VBComponents.Item("UI_QuickToolbarActions")
        $artifactActionCode = $artifactActionModule.CodeModule.Lines(1, $artifactActionModule.CodeModule.CountOfLines)
        $artifactMainModule = $doc.VBProject.VBComponents.Item("UI_QuickToolbar")
        $artifactMainCode = $artifactMainModule.CodeModule.Lines(1, $artifactMainModule.CodeModule.CountOfLines)
        $artifactDispatchCode = $artifactMainCode + "`r`n" + $artifactActionCode
        foreach ($actionCase in $expectedDispatch.GetEnumerator()) {
            Assert-Contains $artifactDispatchCode ('"' + $actionCase.Key + '"') "Built dispatcher should contain $($actionCase.Key)."
        }
    }

    Write-Output "PASS QuickToolbar.Tests"
}
catch {
    throw "Quick toolbar Word COM verification failed. Enable 'Trust access to the VBA project object model' if required. $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)"
}
finally {
    if ($null -ne $doc) {
        $doc.Close($false) | Out-Null
    }
    Stop-RAToolsWordSession -Word $word
}
