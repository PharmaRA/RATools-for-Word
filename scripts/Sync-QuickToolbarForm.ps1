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
$buttonData = Import-RAToolsDataFile (Join-Path $PSScriptRoot "QuickToolbarButtons.psd1")
$toolbarItems = $buttonData.Buttons
$layoutSpec = $buttonData.Layout

$toolbarWidth = [int]$layoutSpec.ToolbarWidth
$toolbarHeight = [int]$layoutSpec.ToolbarHeight
$framelessClientWidth = [int]$layoutSpec.FramelessClientWidth
$headerHeight = [int]$layoutSpec.HeaderHeight
$expandedClientHeight = [int]$layoutSpec.ExpandedClientHeight
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

Private Sub $($item.Name)_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefreshTooltipWatch
End Sub
"@
    }

    $template = @'
Option Explicit

Private Const SETTINGS_APP As String = "RATools"
Private Const SETTINGS_SECTION As String = "QuickToolbar"
Private Const TOOLBAR_WIDTH As Single = __TOOLBAR_WIDTH__
Private Const TOOLBAR_HEIGHT As Single = __TOOLBAR_HEIGHT__
Private Const FRAMELESS_CLIENT_WIDTH As Single = __FRAMELESS_CLIENT_WIDTH__
Private Const HEADER_HEIGHT As Single = __HEADER_HEIGHT__
Private Const EXPANDED_CLIENT_HEIGHT As Single = __EXPANDED_CLIENT_HEIGHT__
Private Const BUTTON_TOP As Single = __BUTTON_TOP__
Private Const BUTTON_WIDTH As Single = __BUTTON_WIDTH__
Private Const BUTTON_HEIGHT As Single = __BUTTON_HEIGHT__
Private Const BUTTON_ROW_STEP As Single = __BUTTON_ROW_STEP__
Private Const BUTTON_COLUMN_GAP As Single = __BUTTON_COLUMN_GAP__
Private Const BUTTON_GROUP_GAP As Single = __BUTTON_GROUP_GAP__

Private m_HoverButtons As Collection
Private m_Dividers As Collection
Private m_ThemeApplied As Boolean
Private m_LastDark As Boolean
Private m_IsCollapsed As Boolean
Private m_IsFrameless As Boolean

Private WithEvents m_HdrBg As MSForms.Label
Private WithEvents m_HdrGrip As MSForms.Label
Private WithEvents m_HdrTitle As MSForms.Label
Private m_HdrDivider As MSForms.Label
Private m_BtnCollapse As clsFlatButton
Private m_BtnClose As clsFlatButton

Private Sub UserForm_Initialize()
    Me.Caption = "RATools"
    Me.Width = TOOLBAR_WIDTH
    Me.Height = TOOLBAR_HEIGHT
    m_IsCollapsed = False
    m_IsFrameless = False
    Set m_HoverButtons = New Collection

    CreateHeaderControls

__CONFIGURE_LINES__

    CreateGroupDividers
    ApplyTheme True
    LayoutToolbarControls
    RestoreToolbarPosition
End Sub

Private Sub UserForm_Activate()
    ApplyTheme True
End Sub

Private Sub CreateHeaderControls()
    Set m_HdrBg = Me.Controls.Add("Forms.Label.1", "hdrBg", True)
    With m_HdrBg
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
    End With

    Set m_HdrGrip = Me.Controls.Add("Forms.Label.1", "hdrGrip", True)
    With m_HdrGrip
        .Caption = ChrW$(&HE784)
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .Font.Name = ThemeIconFontName()
        .Font.Size = 8
        .ControlTipText = "按住拖动悬浮工具栏"
    End With

    Set m_HdrTitle = Me.Controls.Add("Forms.Label.1", "hdrTitle", True)
    With m_HdrTitle
        .Caption = "RATools"
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
        .Font.Bold = True
        .ControlTipText = "按住拖动（双击可折叠/展开）"
    End With

    Set m_HdrDivider = Me.Controls.Add("Forms.Label.1", "hdrDivider", True)
    With m_HdrDivider
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
    End With

    Set m_BtnCollapse = CreateFlatButton(m_HoverButtons, Me, "hdrCollapse", vbNullString, "header", "collapse", Me, 8)
    m_BtnCollapse.SetIcon ChrW$(&HE70D), ThemeIconFontName(), 7.5
    m_BtnCollapse.SetToolTip "折叠工具栏"

    Set m_BtnClose = CreateFlatButton(m_HoverButtons, Me, "hdrClose", vbNullString, "close", "close", Me, 8)
    m_BtnClose.SetIcon ChrW$(&HE8BB), ThemeIconFontName(), 7.5
    m_BtnClose.SetToolTip "隐藏悬浮窗"
End Sub

Public Sub ApplyFramelessMode()
    MakeUserFormWindowFrameless Me.Caption
    m_IsFrameless = True
    ResizeForCurrentState
    LayoutToolbarControls
End Sub

Public Sub ToggleCollapsed()
    m_IsCollapsed = Not m_IsCollapsed
    If Not m_BtnCollapse Is Nothing Then
        If m_IsCollapsed Then
            m_BtnCollapse.SetIcon ChrW$(&HE70E), ThemeIconFontName(), 7.5
            m_BtnCollapse.SetToolTip "展开工具栏"
        Else
            m_BtnCollapse.SetIcon ChrW$(&HE70D), ThemeIconFontName(), 7.5
            m_BtnCollapse.SetToolTip "折叠工具栏"
        End If
    End If
    ResizeForCurrentState
    LayoutToolbarControls
    ReturnFocusToDocumentWindow
End Sub

Private Sub ResizeForCurrentState()
    Dim chromeW As Single
    Dim chromeH As Single
    Dim targetClientH As Single

    If Not m_IsFrameless Then Exit Sub

    On Error Resume Next
    chromeW = Me.Width - Me.InsideWidth
    chromeH = Me.Height - Me.InsideHeight
    If chromeW < 0 Then chromeW = 0
    If chromeH < 0 Then chromeH = 0

    If m_IsCollapsed Then
        targetClientH = HEADER_HEIGHT
    Else
        targetClientH = EXPANDED_CLIENT_HEIGHT
    End If

    Me.Width = FRAMELESS_CLIENT_WIDTH + chromeW
    Me.Height = targetClientH + chromeH
    On Error GoTo 0
End Sub

Private Sub LayoutToolbarControls()
    Dim cw As Single
    Dim ctrl As Object
    Dim divCtrl As MSForms.Label

    On Error Resume Next
    cw = Me.InsideWidth
    If cw <= 0 Then cw = FRAMELESS_CLIENT_WIDTH

    If Not m_HdrBg Is Nothing Then m_HdrBg.Move 0, 0, cw, HEADER_HEIGHT - 0.75
    If Not m_HdrGrip Is Nothing Then m_HdrGrip.Move 4, (HEADER_HEIGHT - 11) / 2, 10, 11
    If Not m_HdrTitle Is Nothing Then m_HdrTitle.Move 15, (HEADER_HEIGHT - 12.5) / 2, cw - 54, 12.5
    If Not m_BtnCollapse Is Nothing Then m_BtnCollapse.MoveTo cw - 38, 2, 17, HEADER_HEIGHT - 4
    If Not m_BtnClose Is Nothing Then m_BtnClose.MoveTo cw - 19, 2, 17, HEADER_HEIGHT - 4
    If Not m_HdrDivider Is Nothing Then
        m_HdrDivider.Move 0, HEADER_HEIGHT - 0.75, cw, 0.75
        m_HdrDivider.Visible = (Not m_IsCollapsed)
    End If

    If Not m_HdrBg Is Nothing Then m_HdrBg.ZOrder 1
    If Not m_HdrGrip Is Nothing Then m_HdrGrip.ZOrder 0
    If Not m_HdrTitle Is Nothing Then m_HdrTitle.ZOrder 0
    If Not m_BtnCollapse Is Nothing Then m_BtnCollapse.MoveTo cw - 38, 2, 17, HEADER_HEIGHT - 4
    If Not m_BtnClose Is Nothing Then m_BtnClose.MoveTo cw - 19, 2, 17, HEADER_HEIGHT - 4
    If Not m_HdrDivider Is Nothing Then m_HdrDivider.ZOrder 0

    For Each ctrl In Me.Controls
        If Left$(ctrl.Name, 3) = "btn" Then
            ctrl.Visible = (Not m_IsCollapsed)
        End If
    Next ctrl

    If Not m_Dividers Is Nothing Then
        For Each divCtrl In m_Dividers
            divCtrl.Visible = (Not m_IsCollapsed)
        Next divCtrl
    End If
    On Error GoTo 0
End Sub

Public Sub ApplyTheme(Optional ByVal force As Boolean = False)
    Dim darkNow As Boolean
    Dim divCtrl As MSForms.Label

    darkNow = IsOfficeDarkTheme()
    If m_ThemeApplied And Not force And (darkNow = m_LastDark) Then Exit Sub

    ApplyFormBaseTheme Me

    On Error Resume Next
    If Not m_HdrBg Is Nothing Then m_HdrBg.BackColor = ThemeCanvasColor()
    If Not m_HdrGrip Is Nothing Then m_HdrGrip.ForeColor = ThemeTextMutedColor()
    If Not m_HdrTitle Is Nothing Then m_HdrTitle.ForeColor = ThemeTextPrimaryColor()
    If Not m_HdrDivider Is Nothing Then StyleDividerLine m_HdrDivider
    On Error GoTo 0

    If Not m_Dividers Is Nothing Then
        For Each divCtrl In m_Dividers
            StyleDividerLine divCtrl
        Next divCtrl
    End If
    RefreshHoverButtons m_HoverButtons
    ApplyFormWindowChrome Me.Caption, False

    m_LastDark = darkNow
    m_ThemeApplied = True
End Sub

Public Sub OnHoverControlChanged(ByVal activeItem As Object)
    ResetHoverButtons m_HoverButtons, activeItem
    RefreshTooltipWatch
End Sub

Public Sub OnFlatButtonClick(ByVal actionKey As String)
    Select Case actionKey
        Case "collapse"
            ToggleCollapsed
        Case "close"
            SaveToolbarPosition
            Me.Hide
    End Select
End Sub

Private Sub m_HdrBg_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    DragUserFormWindow Me.Caption
    SaveToolbarPosition
End Sub

Private Sub m_HdrGrip_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    DragUserFormWindow Me.Caption
    SaveToolbarPosition
End Sub

Private Sub m_HdrTitle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    DragUserFormWindow Me.Caption
    SaveToolbarPosition
End Sub

Private Sub m_HdrBg_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ToggleCollapsed
End Sub

Private Sub m_HdrTitle_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ToggleCollapsed
End Sub

Private Sub m_HdrBg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    RefreshTooltipWatch
End Sub

Private Sub m_HdrGrip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    RefreshTooltipWatch
End Sub

Private Sub m_HdrTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    RefreshTooltipWatch
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    RefreshTooltipWatch
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
        .BackColor = ThemeToolbarTileColor()
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeToolbarTileBorderColor()
        .SpecialEffect = fmSpecialEffectFlat
        .PicturePosition = fmPicturePositionCenter
        .TextAlign = fmTextAlignCenter
        .TabIndex = tabIndex
        .Font.Name = "Microsoft YaHei"
        .Font.Size = 7
    End With

    If HasButtonPicture(buttonItem) Then buttonItem.Caption = vbNullString
    RegisterHoverButton m_HoverButtons, buttonItem, "toolbar", Me
End Sub

Private Sub CreateGroupDividers()
    Dim gapRows As Variant
    Dim idx As Long
    Dim gapRow As Long
    Dim prevBottom As Single
    Dim nextTop As Single
    Dim rowWidth As Single
    Dim divCtrl As MSForms.Label

    Set m_Dividers = New Collection
    gapRows = Array(__GROUP_GAP_LIST__)
    rowWidth = 4 * BUTTON_WIDTH + 3 * BUTTON_COLUMN_GAP

    On Error Resume Next
    For idx = LBound(gapRows) To UBound(gapRows)
        gapRow = CLng(gapRows(idx))
        prevBottom = GetButtonTop(gapRow - 1) + BUTTON_HEIGHT
        nextTop = GetButtonTop(gapRow)
        Set divCtrl = Me.Controls.Add("Forms.Label.1", "divGroup" & CStr(idx + 1), True)
        If Not divCtrl Is Nothing Then
            divCtrl.Move GetButtonLeft(0, 4), prevBottom + (nextTop - prevBottom - 0.75) / 2, rowWidth, 0.75
            StyleDividerLine divCtrl
            m_Dividers.Add divCtrl
        End If
    Next idx
    On Error GoTo 0
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
    ClearHoverButtons m_HoverButtons
    Set m_Dividers = Nothing
    Set m_BtnCollapse = Nothing
    Set m_BtnClose = Nothing
    Set m_HdrDivider = Nothing
    Set m_HdrTitle = Nothing
    Set m_HdrGrip = Nothing
    Set m_HdrBg = Nothing
End Sub

Private Sub UserForm_Terminate()
    SaveToolbarPosition
End Sub
'@

    $groupGapLines = ($layoutSpec.GroupGapRows | ForEach-Object {
        "    If rowIndex >= $_ Then GetButtonTop = GetButtonTop + BUTTON_GROUP_GAP"
    }) -join "`r`n"
    $groupGapList = ($layoutSpec.GroupGapRows -join ", ")

    return $template.Replace("__CONFIGURE_LINES__", ($configureLines -join "`r`n")).Replace(
        "__HANDLERS__",
        ($handlerLines -join "`r`n")
    ).Replace("__TOOLBAR_WIDTH__", [string]$toolbarWidth).Replace(
        "__TOOLBAR_HEIGHT__", [string]$toolbarHeight
    ).Replace("__FRAMELESS_CLIENT_WIDTH__", [string]$framelessClientWidth).Replace(
        "__HEADER_HEIGHT__", [string]$headerHeight
    ).Replace("__EXPANDED_CLIENT_HEIGHT__", [string]$expandedClientHeight).Replace(
        "__BUTTON_TOP__", [string]$buttonTop
    ).Replace("__BUTTON_WIDTH__", [string]$buttonWidth).Replace(
        "__BUTTON_HEIGHT__", [string]$buttonHeight
    ).Replace("__BUTTON_ROW_STEP__", [string]$buttonRowStep).Replace(
        "__BUTTON_COLUMN_GAP__", [string]$buttonColumnGap
    ).Replace("__BUTTON_GROUP_GAP__", [string]$groupGap).Replace(
        "__GROUP_GAP_LINES__", $groupGapLines
    ).Replace("__GROUP_GAP_LIST__", $groupGapList)
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
$themeHelper = $null
$windowHelper = $null
$hoverHelper = $null
$flatBtnHelper = $null

try {
    $word = Start-RAToolsWordSession

    $document = $word.Documents.Add()

    # 导入依赖模块以便 VBE 编译导出
    $themeHelper = $document.VBProject.VBComponents.Import((Join-Path $RepoRoot "modules\UI_Theme.bas"))
    $windowHelper = $document.VBProject.VBComponents.Import((Join-Path $RepoRoot "modules\Core_Window.bas"))
    $hoverHelper = $document.VBProject.VBComponents.Import((Join-Path $RepoRoot "class_modules\clsHoverButton.cls"))
    $flatBtnHelper = $document.VBProject.VBComponents.Import((Join-Path $RepoRoot "class_modules\clsFlatButton.cls"))

    $component = $document.VBProject.VBComponents.Import($formPath)
    $designer = $component.Designer

    $helper = $document.VBProject.VBComponents.Add(1)
    $helper.Name = "QuickToolbarFormSync"
    $helperCode = @'
Option Explicit

Public Sub RunQuickToolbarAction(ByVal actionKey As String)
End Sub

Public Sub RefreshTooltipWatch()
End Sub

Public Sub ReturnFocusToDocumentWindow()
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
    # 剥离 VBE 导出时为 WithEvents 私有变量自动插入的 Attribute 行，保持与模板严格一致
    $exportText = [regex]::Replace(
        $exportText,
        "(?m)^Attribute\s+\w+\.VB_VarHelpID\s*=\s*-1\r?\n",
        ""
    )
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
    Release-ComObjectSafe $themeHelper
    Release-ComObjectSafe $windowHelper
    Release-ComObjectSafe $hoverHelper
    Release-ComObjectSafe $flatBtnHelper
    Release-ComObjectSafe $designer
    Release-ComObjectSafe $component
    Release-ComObjectSafe $document
    Stop-RAToolsWordSession -Word $word
    
    if (Test-Path -LiteralPath $workRoot) {
        Remove-Item -LiteralPath $workRoot -Recurse -Force
    }
}
