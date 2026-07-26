#requires -version 5.1

[CmdletBinding()]
param(
    [string]$RepoRoot
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$formPath = Join-Path $RepoRoot "userforms\frmQuickToolbar.frm"
$frxPath = Join-Path $RepoRoot "userforms\frmQuickToolbar.frx"
$imageRoot = Join-Path $RepoRoot "dotm\customUI\images"
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ("RATools_QuickToolbar_" + [guid]::NewGuid().ToString("N"))

$toolbarItems = @(
    @{ Name = "btnHeading1"; Caption = "编号标题 1"; Tip = "应用标题1-F样式"; Icon = "H1"; Kind = "custom"; Action = "style_heading1" },
    @{ Name = "btnHeading2"; Caption = "编号标题 2"; Tip = "应用标题2-F样式"; Icon = "H2"; Kind = "custom"; Action = "style_heading2" },
    @{ Name = "btnHeading3"; Caption = "编号标题 3"; Tip = "应用标题3-F样式"; Icon = "H3"; Kind = "custom"; Action = "style_heading3" },
    @{ Name = "btnHeading4"; Caption = "编号标题 4"; Tip = "应用标题4-F样式"; Icon = "H4"; Kind = "custom"; Action = "style_heading4" },
    @{ Name = "btnUnnumberedHeading1"; Caption = "无编号标题 1"; Tip = "应用无编号标题1-F样式"; Icon = "UNH1"; Kind = "custom"; Action = "style_unnumbered_heading1" },
    @{ Name = "btnUnnumberedHeading2"; Caption = "无编号标题 2"; Tip = "应用无编号标题2-F样式"; Icon = "UNH2"; Kind = "custom"; Action = "style_unnumbered_heading2" },
    @{ Name = "btnUnnumberedHeading3"; Caption = "无编号标题 3"; Tip = "应用无编号标题3-F样式"; Icon = "UNH3"; Kind = "custom"; Action = "style_unnumbered_heading3" },
    @{ Name = "btnUnnumberedHeading4"; Caption = "无编号标题 4"; Tip = "应用无编号标题4-F样式"; Icon = "UNH4"; Kind = "custom"; Action = "style_unnumbered_heading4" },
    @{ Name = "btnFormatPainter"; Caption = "格式刷"; Tip = "复制并应用所选内容的格式"; Icon = "FormatPainter"; Kind = "mso"; Action = "format_painter" },
    @{ Name = "btnParagraphSettings"; Caption = "段落设置"; Tip = "打开 Word 段落设置对话框"; Icon = "ParagraphDialog"; Kind = "mso"; Action = "paragraph_settings" },
    @{ Name = "btnPageBreakBefore"; Caption = "段前分页"; Tip = "切换选中段落的段前分页属性"; Icon = "PageBreakBefore"; Kind = "custom"; Action = "page_break_before" },
    @{ Name = "btnTableTitle"; Caption = "表标题"; Tip = "应用表标题-F样式"; Icon = "TableTitle"; Kind = "custom"; Action = "style_table_title" },
    @{ Name = "btnFigureTitle"; Caption = "图标题"; Tip = "应用图标题-F样式"; Icon = "FigureTitle"; Kind = "custom"; Action = "style_figure_title" },
    @{ Name = "btnAutoFitTable"; Caption = "根据窗口自动调整表格"; Tip = "将光标所在表格根据窗口宽度自动调整"; Icon = "TableAutoFitWindow"; Kind = "mso-image"; Action = "autofit_table" },
    @{ Name = "btnInsertCrossReference"; Caption = "插入交叉引用"; Tip = "打开插入交叉引用对话框"; Icon = "CrossReferenceInsert"; Kind = "mso"; Action = "insert_cross_reference" },
    @{ Name = "btnUpdateFields"; Caption = "更新域"; Tip = "更新当前选区或光标所在位置的域"; Icon = "FieldsUpdate"; Kind = "mso"; Action = "update_fields" },
    @{ Name = "btnHyperlinksFieldsBlue"; Caption = "超链接和域批量设置为蓝色"; Tip = "运行超链接和域批量设置为蓝色宏"; Icon = "HyperlinkInsert"; Kind = "mso"; Action = "hyperlinks_fields_blue" },
    @{ Name = "btnAcceptRevisionsComments"; Caption = "接受修订并删除批注"; Tip = "运行接受修订并删除批注宏"; Icon = "ReviewAcceptChange"; Kind = "mso"; Action = "accept_revisions_comments" },
    @{ Name = "btnDetectHighlights"; Caption = "检测高亮内容"; Tip = "运行检测高亮内容宏"; Icon = "TextHighlightColorPicker"; Kind = "mso"; Action = "detect_highlights" }
)

$toolbarWidth = 134
$toolbarHeight = 168
$buttonWidth = 26
$buttonHeight = 24
$buttonRowStep = 25
$buttonColumnGap = 3
$buttonTop = 3
$groupGap = 4
$pictureSize = 20

function Release-ComObjectSafe {
    param($ComObject)

    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Get-ButtonRow {
    param([int]$Index)

    if ($Index -lt 4) { return 0 }
    if ($Index -lt 8) { return 1 }
    if ($Index -lt 12) { return 2 }
    if ($Index -lt 16) { return 3 }
    return 4
}

function Get-ButtonColumn {
    param([int]$Index)

    if ($Index -lt 16) { return $Index % 4 }
    return $Index - 16
}

function Get-ButtonColumnCount {
    param([int]$Index)

    if ($Index -lt 16) { return 4 }
    return 3
}

function Get-ButtonTop {
    param([int]$Row)

    $top = $buttonTop + ($Row * $buttonRowStep)
    if ($Row -ge 2) { $top += $groupGap }
    if ($Row -ge 4) { $top += $groupGap }
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
        $row = Get-ButtonRow -Index $i
        $column = Get-ButtonColumn -Index $i
        $columnCount = Get-ButtonColumnCount -Index $i
        "    ConfigureButton Me.$($toolbarItems[$i].Name), $row, $column, $columnCount, $i"
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
Private Const TOOLBAR_WIDTH As Single = 134
Private Const TOOLBAR_HEIGHT As Single = 168
Private Const BUTTON_TOP As Single = 3
Private Const BUTTON_WIDTH As Single = 26
Private Const BUTTON_HEIGHT As Single = 24
Private Const BUTTON_ROW_STEP As Single = 25
Private Const BUTTON_COLUMN_GAP As Single = 3
Private Const BUTTON_GROUP_GAP As Single = 4

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
    If rowIndex >= 2 Then GetButtonTop = GetButtonTop + BUTTON_GROUP_GAP
    If rowIndex >= 4 Then GetButtonTop = GetButtonTop + BUTTON_GROUP_GAP
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

    return $template.Replace("__CONFIGURE_LINES__", ($configureLines -join "`r`n")).Replace(
        "__HANDLERS__",
        ($handlerLines -join "`r`n")
    )
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
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1
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
        $row = Get-ButtonRow -Index $i
        $column = Get-ButtonColumn -Index $i
        $columnCount = Get-ButtonColumnCount -Index $i
        $control = $designer.Controls.Add("Forms.Label.1", $item.Name, $true)
        $control.Caption = $item.Caption
        $control.ControlTipText = $item.Caption + "：" + $item.Tip
        $control.Tag = if ($item.Kind -eq "mso-image") { "imageMso:" + $item.Icon } elseif ($item.Kind -eq "mso") { "idMso:" + $item.Icon } else { $item.Icon }
        $control.Left = [single](Get-ButtonLeft -Column $column -ColumnCount $columnCount)
        $control.Top = [single](Get-ButtonTop -Row $row)
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
    if ($null -ne $word) { $word.Quit() | Out-Null }

    Release-ComObjectSafe $helper
    Release-ComObjectSafe $designer
    Release-ComObjectSafe $component
    Release-ComObjectSafe $document
    Release-ComObjectSafe $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if (Test-Path -LiteralPath $workRoot) {
        Remove-Item -LiteralPath $workRoot -Recurse -Force
    }
}
