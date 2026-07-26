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

function Get-ExpectedQuickToolbarLayout {
    param([int]$Index)

    if ($Index -lt 4) {
        $row = 0
    }
    elseif ($Index -lt 8) {
        $row = 1
    }
    elseif ($Index -lt 12) {
        $row = 2
    }
    elseif ($Index -lt 16) {
        $row = 3
    }
    else {
        $row = 4
    }

    if ($Index -lt 16) {
        $column = $Index % 4
        $columnCount = 4
    }
    else {
        $column = $Index - 16
        $columnCount = 3
    }

    return [pscustomobject]@{
        Row = $row
        Column = $column
        ColumnCount = $columnCount
    }
}

$expectedItems = @(
    [pscustomobject]@{ Name = "btnHeading1"; Caption = "编号标题 1"; Tip = "应用标题1-F样式"; Tag = "H1"; Action = "style_heading1" },
    [pscustomobject]@{ Name = "btnHeading2"; Caption = "编号标题 2"; Tip = "应用标题2-F样式"; Tag = "H2"; Action = "style_heading2" },
    [pscustomobject]@{ Name = "btnHeading3"; Caption = "编号标题 3"; Tip = "应用标题3-F样式"; Tag = "H3"; Action = "style_heading3" },
    [pscustomobject]@{ Name = "btnHeading4"; Caption = "编号标题 4"; Tip = "应用标题4-F样式"; Tag = "H4"; Action = "style_heading4" },
    [pscustomobject]@{ Name = "btnUnnumberedHeading1"; Caption = "无编号标题 1"; Tip = "应用无编号标题1-F样式"; Tag = "UNH1"; Action = "style_unnumbered_heading1" },
    [pscustomobject]@{ Name = "btnUnnumberedHeading2"; Caption = "无编号标题 2"; Tip = "应用无编号标题2-F样式"; Tag = "UNH2"; Action = "style_unnumbered_heading2" },
    [pscustomobject]@{ Name = "btnUnnumberedHeading3"; Caption = "无编号标题 3"; Tip = "应用无编号标题3-F样式"; Tag = "UNH3"; Action = "style_unnumbered_heading3" },
    [pscustomobject]@{ Name = "btnUnnumberedHeading4"; Caption = "无编号标题 4"; Tip = "应用无编号标题4-F样式"; Tag = "UNH4"; Action = "style_unnumbered_heading4" },
    [pscustomobject]@{ Name = "btnFormatPainter"; Caption = "格式刷"; Tip = "复制并应用所选内容的格式"; Tag = "idMso:FormatPainter"; Action = "format_painter" },
    [pscustomobject]@{ Name = "btnParagraphSettings"; Caption = "段落设置"; Tip = "打开 Word 段落设置对话框"; Tag = "idMso:ParagraphDialog"; Action = "paragraph_settings" },
    [pscustomobject]@{ Name = "btnPageBreakBefore"; Caption = "段前分页"; Tip = "切换选中段落的段前分页属性"; Tag = "PageBreakBefore"; Action = "page_break_before" },
    [pscustomobject]@{ Name = "btnTableTitle"; Caption = "表标题"; Tip = "应用表标题-F样式"; Tag = "TableTitle"; Action = "style_table_title" },
    [pscustomobject]@{ Name = "btnFigureTitle"; Caption = "图标题"; Tip = "应用图标题-F样式"; Tag = "FigureTitle"; Action = "style_figure_title" },
    [pscustomobject]@{ Name = "btnAutoFitTable"; Caption = "根据窗口自动调整表格"; Tip = "将光标所在表格根据窗口宽度自动调整"; Tag = "imageMso:TableAutoFitWindow"; Action = "autofit_table" },
    [pscustomobject]@{ Name = "btnInsertCrossReference"; Caption = "插入交叉引用"; Tip = "打开插入交叉引用对话框"; Tag = "idMso:CrossReferenceInsert"; Action = "insert_cross_reference" },
    [pscustomobject]@{ Name = "btnUpdateFields"; Caption = "更新域"; Tip = "更新当前选区或光标所在位置的域"; Tag = "idMso:FieldsUpdate"; Action = "update_fields" },
    [pscustomobject]@{ Name = "btnHyperlinksFieldsBlue"; Caption = "超链接和域批量设置为蓝色"; Tip = "运行超链接和域批量设置为蓝色宏"; Tag = "idMso:HyperlinkInsert"; Action = "hyperlinks_fields_blue" },
    [pscustomobject]@{ Name = "btnAcceptRevisionsComments"; Caption = "接受修订并删除批注"; Tip = "运行接受修订并删除批注宏"; Tag = "idMso:ReviewAcceptChange"; Action = "accept_revisions_comments" },
    [pscustomobject]@{ Name = "btnDetectHighlights"; Caption = "检测高亮内容"; Tip = "运行检测高亮内容宏"; Tag = "idMso:TextHighlightColorPicker"; Action = "detect_highlights" }
)

$expectedDispatch = [ordered]@{
    style_heading1            = "style:标题1-F"
    style_heading2            = "style:标题2-F"
    style_heading3            = "style:标题3-F"
    style_heading4            = "style:标题4-F"
    style_unnumbered_heading1 = "style:无编号标题1-F"
    style_unnumbered_heading2 = "style:无编号标题2-F"
    style_unnumbered_heading3 = "style:无编号标题3-F"
    style_unnumbered_heading4 = "style:无编号标题4-F"
    format_painter            = "mso:FormatPainter"
    paragraph_settings        = "mso:ParagraphDialog"
    page_break_before         = "page_break_before"
    style_table_title         = "style:表标题-F"
    style_figure_title        = "style:图标题-F"
    autofit_table             = "autofit_table"
    insert_cross_reference    = "mso:CrossReferenceInsert"
    update_fields             = "mso:FieldsUpdate"
    hyperlinks_fields_blue    = "hyperlinks_fields_blue"
    accept_revisions_comments = "accept_revisions_comments"
    detect_highlights         = "detect_highlights"
}

Write-Output "Running QuickToolbar source checks"
$formPath = Join-Path $repoRoot "userforms\frmQuickToolbar.frm"
$frxPath = Join-Path $repoRoot "userforms\frmQuickToolbar.frx"
$modulePath = Join-Path $repoRoot "modules\Mod_QuickToolbar.bas"
$actionModulePath = Join-Path $repoRoot "modules\Mod_QuickToolbarActions.bas"
$syncScriptPath = Join-Path $repoRoot "scripts\Sync-QuickToolbarForm.ps1"
$ribbonPath = Join-Path $repoRoot "dotm\customUI\customUI14.xml"

foreach ($requiredPath in @($formPath, $frxPath, $modulePath, $actionModulePath, $syncScriptPath, $ribbonPath)) {
    Assert-True (Test-Path -LiteralPath $requiredPath -PathType Leaf) "Required quick toolbar file should exist: $requiredPath"
}
Assert-True ((Get-Item -LiteralPath $frxPath).Length -gt 25000) "Quick toolbar FRX should contain all 19 embedded icon pictures."

$formText = Read-VbaSource -Path $formPath
$moduleText = Read-VbaSource -Path $modulePath
$actionModuleText = Read-VbaSource -Path $actionModulePath
$syncScriptText = Get-Content -LiteralPath $syncScriptPath -Raw -Encoding UTF8
$ribbonText = Get-Content -LiteralPath $ribbonPath -Raw

Assert-Contains $formText 'OleObjectBlob   =   "frmQuickToolbar.frx":0000' "Form should reference its FRX resource."
Assert-Contains $formText "ShowModal       =   0" "Quick toolbar should be modeless."
Assert-Contains $formText 'Me.Caption = "RATools"' "Quick toolbar should use a short readable title."
Assert-Contains $formText "Private Const TOOLBAR_WIDTH As Single = 134" "Toolbar should fit its four-column grid."
Assert-Contains $formText "Private Const TOOLBAR_HEIGHT As Single = 168" "Toolbar should fit its compact five-row grid."
Assert-Contains $formText "Private Const BUTTON_WIDTH As Single = 26" "Toolbar should use compact icon tiles."
Assert-Contains $formText "Private Const BUTTON_HEIGHT As Single = 24" "Toolbar should use compact icon tiles."
Assert-Contains $formText "Private Const BUTTON_ROW_STEP As Single = 25" "Toolbar should use compact row spacing."
Assert-Contains $formText "Private Const BUTTON_COLUMN_GAP As Single = 3" "Toolbar should use compact column spacing."
Assert-Contains $formText "If rowIndex >= 2 Then" "Toolbar should separate the heading and command groups."
Assert-Contains $formText "If rowIndex >= 4 Then" "Toolbar should separate the command and macro groups."
Assert-Contains $formText "GetButtonLeft = (Me.InsideWidth - rowWidth) / 2" "Each button row should be centered from the live client width."
Assert-Contains $formText "ByVal buttonItem As MSForms.Label" "Toolbar controls should use flat label tiles."
Assert-Contains $formText ".PicturePosition = fmPicturePositionCenter" "Toolbar icons should be centered inside each tile."
Assert-Contains $moduleText "Public Sub ToggleQuickToolbar" "Ribbon toggle callback should exist."
Assert-Contains $moduleText "Public Sub RunQuickToolbarAction" "Action dispatcher should exist."
Assert-Contains $moduleText "TryRunExpandedQuickToolbarAction(actionKey)" "Dispatcher should delegate the expanded button set."
Assert-Contains $actionModuleText "Public Function TryRunExpandedQuickToolbarAction" "Expanded action dispatcher should exist."
Assert-Contains $actionModuleText 'Application.CommandBars.ExecuteMso commandId' "Built-in Word commands should use ExecuteMso."
Assert-Contains $ribbonText 'onAction="ToggleQuickToolbar"' "Ribbon should expose the quick toolbar callback."

for ($i = 0; $i -lt $expectedItems.Count; $i++) {
    $item = $expectedItems[$i]
    $layout = Get-ExpectedQuickToolbarLayout -Index $i
    $expectedConfigure = "ConfigureButton Me.$($item.Name), $($layout.Row), $($layout.Column), $($layout.ColumnCount), $i"
    Assert-Contains $formText $expectedConfigure "Missing grid layout wiring for $($item.Name)."
    Assert-Contains $formText ("Private Sub " + $item.Name + "_Click()") "Missing click handler for $($item.Name)."
    Assert-Contains $formText ('RunQuickToolbarAction "' + $item.Action + '"') "Wrong action wiring for $($item.Name)."
    Assert-Contains $syncScriptText ('Name = "' + $item.Name + '"') "Generator should define $($item.Name)."
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

    foreach ($sourceFile in @($modulePath, $actionModulePath, $formPath)) {
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
    Assert-True ($buttonCount -eq $expectedItems.Count) "Quick toolbar should contain 19 controls; got $buttonCount."

    $toolbarCaption = [string]$word.Run("GetQuickToolbarCaptionForTest")
    Assert-True ($toolbarCaption -eq "RATools") "Quick toolbar title should be readable; got $toolbarCaption."

    $formSize = ([string]$word.Run("GetQuickToolbarFormSizeForTest")).Split("|")
    Assert-True ($formSize.Count -eq 4) "Could not inspect quick toolbar runtime dimensions."
    $formWidth = [double]$formSize[0]
    $formHeight = [double]$formSize[1]
    $insideWidth = [double]$formSize[2]
    $insideHeight = [double]$formSize[3]
    Assert-True ([Math]::Abs($formWidth - 134) -lt 0.5) "Toolbar runtime width should be 134pt; got $formWidth."
    Assert-True ([Math]::Abs($formHeight - 168) -lt 0.5) "Toolbar runtime height should be 168pt; got $formHeight."
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
        Assert-True ([Math]::Abs($width - 26) -lt 0.1) "$controlName should be 26pt wide; got $width."
        Assert-True ([Math]::Abs($height - 24) -lt 0.1) "$controlName should be 24pt high; got $height."

        $rowWidth = ($layout.ColumnCount * 26) + (($layout.ColumnCount - 1) * 3)
        $expectedLeft = (($insideWidth - $rowWidth) / 2) + ($layout.Column * 29)
        $expectedTop = 3 + ($layout.Row * 25)
        if ($layout.Row -ge 2) { $expectedTop += 4 }
        if ($layout.Row -ge 4) { $expectedTop += 4 }
        Assert-True ([Math]::Abs($left - $expectedLeft) -lt 0.15) "$controlName has an unexpected horizontal grid position: $left."
        Assert-True ([Math]::Abs($top - $expectedTop) -lt 0.15) "$controlName has an unexpected row position: $top."
        Assert-True ($left -ge -0.15) "$controlName should fit inside the left edge of the form."
        Assert-True (($left + $width) -le ($insideWidth + 0.15)) "$controlName should fit inside the right edge of the form."
        Assert-True (($top + $height) -le ($insideHeight + 0.15)) "$controlName should fit fully inside the form; bottom=$($top + $height), insideHeight=$insideHeight."
    }

    [void]$word.Run("ReleaseQuickToolbarForTest")
    $toolbarComponent = $doc.VBProject.VBComponents.Item("frmQuickToolbar")
    Assert-True ([int]$toolbarComponent.Designer.Controls.Count -eq $expectedItems.Count) "Designer should contain exactly 19 controls."
    for ($i = 0; $i -lt $expectedItems.Count; $i++) {
        $item = $expectedItems[$i]
        $control = $toolbarComponent.Designer.Controls.Item([string]$item.Name)
        $layout = Get-ExpectedQuickToolbarLayout -Index $i
        $rowWidth = ($layout.ColumnCount * 26) + (($layout.ColumnCount - 1) * 3)
        $expectedDesignerLeft = ((134 - $rowWidth) / 2) + ($layout.Column * 29)
        $expectedDesignerTop = 3 + ($layout.Row * 25)
        if ($layout.Row -ge 2) { $expectedDesignerTop += 4 }
        if ($layout.Row -ge 4) { $expectedDesignerTop += 4 }
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
        Assert-True ($componentNames -contains "Mod_QuickToolbarActions") "Built dotm should contain Mod_QuickToolbarActions."
        Assert-True ($componentNames -contains "frmQuickToolbar") "Built dotm should contain frmQuickToolbar."

        $artifactToolbar = $doc.VBProject.VBComponents.Item("frmQuickToolbar")
        Assert-True ([int]$artifactToolbar.Designer.Controls.Count -eq $expectedItems.Count) "Built toolbar should contain exactly 19 controls."
        $artifactCode = $artifactToolbar.CodeModule.Lines(1, $artifactToolbar.CodeModule.CountOfLines)
        Assert-Contains $artifactCode "GetButtonLeft = (Me.InsideWidth - rowWidth) / 2" "Built toolbar should preserve row centering."
        Assert-Contains $artifactCode "If rowIndex >= 2 Then" "Built toolbar should preserve the first group gap."
        Assert-Contains $artifactCode "If rowIndex >= 4 Then" "Built toolbar should preserve the second group gap."

        for ($i = 0; $i -lt $expectedItems.Count; $i++) {
            $item = $expectedItems[$i]
            $control = $artifactToolbar.Designer.Controls.Item([string]$item.Name)
            $layout = Get-ExpectedQuickToolbarLayout -Index $i
            $rowWidth = ($layout.ColumnCount * 26) + (($layout.ColumnCount - 1) * 3)
            $expectedArtifactLeft = ((134 - $rowWidth) / 2) + ($layout.Column * 29)
            $expectedArtifactTop = 3 + ($layout.Row * 25)
            if ($layout.Row -ge 2) { $expectedArtifactTop += 4 }
            if ($layout.Row -ge 4) { $expectedArtifactTop += 4 }
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

        $artifactActionModule = $doc.VBProject.VBComponents.Item("Mod_QuickToolbarActions")
        $artifactActionCode = $artifactActionModule.CodeModule.Lines(1, $artifactActionModule.CodeModule.CountOfLines)
        $artifactMainModule = $doc.VBProject.VBComponents.Item("Mod_QuickToolbar")
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
    if ($null -ne $word) {
        $word.Quit() | Out-Null
    }
}
