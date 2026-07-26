#requires -version 5.1

[CmdletBinding()]
param(
    [string]$RepoRoot
)

$ErrorActionPreference = "Stop"

Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$formPath = Join-Path $RepoRoot "userforms\frmQuickToolbar.frm"
$frxPath = Join-Path $RepoRoot "userforms\frmQuickToolbar.frx"
$imageRoot = Join-Path $RepoRoot "dotm\customUI\images"
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ("RATools_QuickToolbar_" + [guid]::NewGuid().ToString("N"))

# 按钮与布局定义单一事实源：QuickToolbarButtons.psd1
# （tests/QuickToolbar.Tests.ps1 从同一文件推导期望值）
$buttonData = Import-PowerShellDataFile (Join-Path $PSScriptRoot "QuickToolbarButtons.psd1")
$toolbarItems = $buttonData.Buttons
$layoutSpec = $buttonData.Layout

$toolbarWidth = [int]$layoutSpec.ToolbarWidth
$toolbarHeight = [int]$layoutSpec.ToolbarHeight
$buttonWidth = [int]$layoutSpec.ButtonWidth
$buttonHeight = [int]$layoutSpec.ButtonHeight
$buttonRowStep = [int]$layoutSpec.RowStep
$buttonColumnGap = [int]$layoutSpec.ColumnGap
$buttonTop = [int]$layoutSpec.ButtonTop
$groupGap = [int]$layoutSpec.GroupGap
$pictureSize = [int]$layoutSpec.PictureSize

function Release-ComObjectSafe {
    param($ComObject)

    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

# 行/列/列数直接取自 psd1 按钮定义；Top/Left 由布局参数推导
function Get-ButtonTop {
    param([int]$Row)

    $top = $buttonTop + ($Row * $buttonRowStep)
    foreach ($gapRow in $layoutSpec.GroupGapRows) {
        if ($Row -ge $gapRow) { $top += $groupGap }
    }
    return $top
}

function Get-ButtonLeft {
    param(
        [int]$Column,
        [int]$ColumnCount
    )

    $rowWidth = ($ColumnCount * $buttonWidth) + (($ColumnCount - 1) * $buttonColumnGap)
    return (($toolbarWidth - $rowWidth) / 2) + ($Column * ($buttonWidth + $buttonColumnGap))
}

function Convert-RibbonImageToBmp {
    param(
        [Parameter(Mandatory)] [string]$SourcePath,
        [Parameter(Mandatory)] [string]$DestinationPath
    )

    $source = [Drawing.Image]::FromFile($SourcePath)
    $bitmap = New-Object Drawing.Bitmap $pictureSize, $pictureSize
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([Drawing.Color]::White)
        $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $scale = [Math]::Min($pictureSize / $source.Width, $pictureSize / $source.Height)
        $width = [Math]::Max(1, [int][Math]::Round($source.Width * $scale))
        $height = [Math]::Max(1, [int][Math]::Round($source.Height * $scale))
        $left = [int](($pictureSize - $width) / 2)
        $top = [int](($pictureSize - $height) / 2)
        $graphics.DrawImage($source, $left, $top, $width, $height)
        $bitmap.Save($DestinationPath, [Drawing.Imaging.ImageFormat]::Bmp)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
        $source.Dispose()
    }
}

function New-FormCode {
    $configureLines = for ($i = 0; $i -lt $toolbarItems.Count; $i++) {
        $item = $toolbarItems[$i]
        "    ConfigureButton Me.$($item.Name), $($item.Row), $($item.Column), $($item.ColumnCount), $i"
    }

    $handlerLines = foreach ($item in $toolbarItems) {
        @"
Private Sub $($item.Name)_Click()
    RunQuickToolbarAction "$($item.Action)"
End Sub
"@
    }

    $template = @'
Option Explicit

Private Const SETTINGS_APP As String = "RATools"
Private Const SETTINGS_SECTION As String = "QuickToolbar"
Private Const TOOLBAR_WIDTH As Single = __TOOLBAR_WIDTH__
Private Const TOOLBAR_HEIGHT As Single = __TOOLBAR_HEIGHT__
Private Const BUTTON_TOP As Single = __BUTTON_TOP__
Private Const BUTTON_WIDTH As Single = __BUTTON_WIDTH__
Private Const BUTTON_HEIGHT As Single = __BUTTON_HEIGHT__
Private Const BUTTON_ROW_STEP As Single = __BUTTON_ROW_STEP__
Private Const BUTTON_COLUMN_GAP As Single = __BUTTON_COLUMN_GAP__
Private Const BUTTON_GROUP_GAP As Single = __BUTTON_GROUP_GAP__

Private Sub UserForm_Initialize()
    Me.Caption = "RATools"
    Me.Width = TOOLBAR_WIDTH
    Me.Height = TOOLBAR_HEIGHT
    Me.BackColor = RGB(247, 248, 250)

__CONFIGURE_LINES__

    RestoreToolbarPosition
End Sub

Private Sub ConfigureButton(ByVal buttonItem As MSForms.Label, _
                            ByVal rowIndex As Long, _
                            ByVal columnIndex As Long, _
                            ByVal columnCount As Long, _
                            ByVal tabIndex As Long)
    With buttonItem
        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT
        .Left = GetButtonLeft(columnIndex, columnCount)
        .Top = GetButtonTop(rowIndex)
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(221, 226, 232)
        .PicturePosition = fmPicturePositionCenter
        .TextAlign = fmTextAlignCenter
        .TabIndex = tabIndex
        .Font.Name = "Microsoft YaHei"
        .Font.Size = 7
    End With

    If HasButtonPicture(buttonItem) Then buttonItem.Caption = vbNullString
End Sub

Private Function GetButtonLeft(ByVal columnIndex As Long, _
                               ByVal columnCount As Long) As Single
    Dim rowWidth As Single

    rowWidth = columnCount * BUTTON_WIDTH + _
               (columnCount - 1) * BUTTON_COLUMN_GAP
    GetButtonLeft = (Me.InsideWidth - rowWidth) / 2 + _
                    columnIndex * (BUTTON_WIDTH + BUTTON_COLUMN_GAP)
End Function

Private Function GetButtonTop(ByVal rowIndex As Long) As Single
    GetButtonTop = BUTTON_TOP + rowIndex * BUTTON_ROW_STEP
__GROUP_GAP_LINES__
End Function

Private Function HasButtonPicture(ByVal buttonItem As MSForms.Label) As Boolean
    Dim pictureItem As Object

    On Error GoTo ErrH
    Set pictureItem = buttonItem.Picture
    HasButtonPicture = Not (pictureItem Is Nothing)
    Exit Function

ErrH:
    HasButtonPicture = False
End Function

__HANDLERS__

Private Sub RestoreToolbarPosition()
    Dim savedLeft As String
    Dim savedTop As String

    savedLeft = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Left", "")
    savedTop = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Top", "")

    If IsNumeric(savedLeft) And IsNumeric(savedTop) Then
        Me.Left = CSng(savedLeft)
        Me.Top = CSng(savedTop)
    Else
        PositionNearWordWindow
    End If
End Sub

Private Sub PositionNearWordWindow()
    On Error GoTo UseFallback

    Me.Left = Application.Left + Application.Width - Me.Width - 24
    Me.Top = Application.Top + 55
    Exit Sub

UseFallback:
    Me.Left = 80
    Me.Top = 55
End Sub

Public Sub SaveToolbarPosition()
    On Error Resume Next
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Left", CStr(Me.Left)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Top", CStr(Me.Top)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveToolbarPosition
End Sub

Private Sub UserForm_Terminate()
    SaveToolbarPosition
End Sub
'@

    $groupGapLines = ($layoutSpec.GroupGapRows | ForEach-Object {
        "    If rowIndex >= $_ Then GetButtonTop = GetButtonTop + BUTTON_GROUP_GAP"
    }) -join "`r`n"

    return $template.Replace("__CONFIGURE_LINES__", ($configureLines -join "`r`n")).Replace(
        "__HANDLERS__",
        ($handlerLines -join "`r`n")
    ).Replace("__TOOLBAR_WIDTH__", [string]$toolbarWidth).Replace(
        "__TOOLBAR_HEIGHT__", [string]$toolbarHeight
    ).Replace("__BUTTON_TOP__", [string]$buttonTop).Replace(
        "__BUTTON_WIDTH__", [string]$buttonWidth
    ).Replace("__BUTTON_HEIGHT__", [string]$buttonHeight).Replace(
        "__BUTTON_ROW_STEP__", [string]$buttonRowStep
    ).Replace("__BUTTON_COLUMN_GAP__", [string]$buttonColumnGap).Replace(
        "__BUTTON_GROUP_GAP__", [string]$groupGap
    ).Replace("__GROUP_GAP_LINES__", $groupGapLines)
}

Add-Type -AssemblyName System.Drawing
[void](New-Item -ItemType Directory -Path $workRoot)

$pictureCodeLines = New-Object System.Collections.Generic.List[string]
foreach ($item in $toolbarItems) {
    $pictureCodeLines.Add('    currentName = "' + $item.Name + '"')
    if ($item.Kind -eq "custom") {
        $sourcePath = Join-Path $imageRoot ($item.Icon + ".png")
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw "Missing Ribbon image: $sourcePath"
        }

        $bmpPath = Join-Path $workRoot ($item.Name + ".bmp")
        Convert-RibbonImageToBmp -SourcePath $sourcePath -DestinationPath $bmpPath
        $escapedBmpPath = $bmpPath.Replace('"', '""')
        $pictureCodeLines.Add('    Set ThisDocument.VBProject.VBComponents("frmQuickToolbar").Designer.Controls("' +
            $item.Name + '").Picture = LoadPicture("' + $escapedBmpPath + '")')
    }
    else {
        $pictureCodeLines.Add('    Set ThisDocument.VBProject.VBComponents("frmQuickToolbar").Designer.Controls("' +
            $item.Name + '").Picture = Application.CommandBars.GetImageMso("' + $item.Icon + '", ' +
            $pictureSize + ', ' + $pictureSize + ')')
    }
}

$word = $null
$document = $null
$component = $null
$designer = $null
$helper = $null

try {
    $word = Start-RAToolsWordSession

    $document = $word.Documents.Add()

    $component = $document.VBProject.VBComponents.Import($formPath)
    $designer = $component.Designer

    $helper = $document.VBProject.VBComponents.Add(1)
    $helper.Name = "QuickToolbarFormSync"
    $helperCode = @'
Option Explicit

Public Sub RunQuickToolbarAction(ByVal actionKey As String)
End Sub

Public Function SetQuickToolbarPictures() As String
    Dim currentName As String

    On Error GoTo ErrH
__PICTURE_ASSIGNMENTS__
    Exit Function

ErrH:
    SetQuickToolbarPictures = currentName & ": " & Err.Description
End Function
'@
    $helperCode = $helperCode.Replace("__PICTURE_ASSIGNMENTS__", ($pictureCodeLines -join "`r`n"))
    $helper.CodeModule.AddFromString($helperCode)

    $existingControlNames = @()
    foreach ($control in $designer.Controls) {
        $existingControlNames += [string]$control.Name
    }
    foreach ($controlName in $existingControlNames) {
        $designer.Controls.Remove($controlName)
    }

    $component.Properties.Item("Caption").Value = "RATools"
    $component.Properties.Item("Width").Value = $toolbarWidth
    $component.Properties.Item("Height").Value = $toolbarHeight

    for ($i = 0; $i -lt $toolbarItems.Count; $i++) {
        $item = $toolbarItems[$i]
        $control = $designer.Controls.Add("Forms.Label.1", $item.Name, $true)
        $control.Caption = $item.Caption
        $control.ControlTipText = $item.Caption + "：" + $item.Tip
        $control.Tag = if ($item.Kind -eq "mso-image") { "imageMso:" + $item.Icon } elseif ($item.Kind -eq "mso") { "idMso:" + $item.Icon } else { $item.Icon }
        $control.Left = [single](Get-ButtonLeft -Column $item.Column -ColumnCount $item.ColumnCount)
        $control.Top = [single](Get-ButtonTop -Row $item.Row)
        $control.Width = [single]$buttonWidth
        $control.Height = [single]$buttonHeight
        $control.BackStyle = 1
        $control.BackColor = 16777215
        $control.BorderStyle = 1
        $control.BorderColor = 15262429
        $control.SpecialEffect = 0
        $control.TextAlign = 2
        $control.TabIndex = $i
    }

    $pictureError = [string]$word.Run("SetQuickToolbarPictures")
    if (-not [string]::IsNullOrWhiteSpace($pictureError)) {
        throw "Could not embed quick toolbar icon for $pictureError"
    }

    $codeModule = $component.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString((New-FormCode))

    $exportPath = Join-Path $workRoot "frmQuickToolbar.frm"
    $component.Export($exportPath)
    $exportFrxPath = [IO.Path]::ChangeExtension($exportPath, ".frx")
    if (-not (Test-Path -LiteralPath $exportFrxPath -PathType Leaf)) {
        throw "The exported quick toolbar did not include an FRX resource."
    }

    $vbaEncoding = [Text.Encoding]::GetEncoding(936)
    $exportText = [IO.File]::ReadAllText($exportPath, $vbaEncoding)
    $exportText = [regex]::Replace(
        $exportText,
        "[ `t]+(?=`r?$)",
        "",
        [Text.RegularExpressions.RegexOptions]::Multiline
    ).TrimEnd("`r", "`n") + "`r`n"
    [IO.File]::WriteAllText($exportPath, $exportText, $vbaEncoding)

    Copy-Item -LiteralPath $exportPath -Destination $formPath -Force
    Copy-Item -LiteralPath $exportFrxPath -Destination $frxPath -Force
    Write-Output "Synchronized $($toolbarItems.Count) quick toolbar controls."
}
finally {
    if ($null -ne $document) { $document.Close($false) | Out-Null }

    Release-ComObjectSafe $helper
    Release-ComObjectSafe $designer
    Release-ComObjectSafe $component
    Release-ComObjectSafe $document
    Stop-RAToolsWordSession -Word $word
    
    if (Test-Path -LiteralPath $workRoot) {
        Remove-Item -LiteralPath $workRoot -Recurse -Force
    }
}
