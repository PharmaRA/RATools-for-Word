$ErrorActionPreference = "Stop"

# UI 主题与窗体美化静态契约（无 COM）：
# 1) 所有新建/改造 VBA 源码严格满足 GBK(936) 无损编码 + CRLF；
# 2) UI_Theme.bas 正确实现随 Word 浅/深主题自动切换（Word Fluent 浅色 + 石墨深色）；
# 3) Core_Window.bas、clsHoverButton.cls、clsFlatButton.cls、clsMacroListRow.cls 暴露原语；
# 4) 全部 6 个 UserForm 接线主题与 DWM 窗口外观，Sync-QuickToolbarForm 与 frmQuickToolbar 保持同步；
# 5) frmBatchMode 资源齐全且 Core_UI.bas 保留无窗体环境下的 InputBox 回退。

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Import-Module (Join-Path $repoRoot "scripts\RATools.Build.psm1") -Force

$assertions = 0

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    $script:assertions++
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
    param([string]$RelativePath)

    $fullPath = Join-Path $repoRoot $RelativePath
    Assert-True (Test-Path -LiteralPath $fullPath -PathType Leaf) "Required VBA source should exist: $RelativePath"

    $gbk = [System.Text.Encoding]::GetEncoding(936)
    $bytes = [System.IO.File]::ReadAllBytes($fullPath)
    $decoded = $gbk.GetString($bytes)
    $roundTrip = $gbk.GetBytes($decoded)

    Assert-True ($roundTrip.Length -eq $bytes.Length) "$RelativePath must be lossless GBK(936)."
    Assert-True (-not $decoded.Contains([char]0xFFFD)) "$RelativePath must not contain replacement characters."
    Assert-True (-not $decoded.Contains("`n`n`n")) "$RelativePath must not contain stray LF blank-line runs."
    Assert-True ($decoded.Contains("`r`n")) "$RelativePath must use CRLF line endings."
    return $decoded
}

Write-Output "Running UITheme static checks"

$windowText = Read-VbaSource "modules\Core_Window.bas"
$themeText = Read-VbaSource "modules\UI_Theme.bas"
$coreUiText = Read-VbaSource "modules\Core_UI.bas"
$qtModuleText = Read-VbaSource "modules\UI_QuickToolbar.bas"
$hoverClassText = Read-VbaSource "class_modules\clsHoverButton.cls"
$flatBtnClassText = Read-VbaSource "class_modules\clsFlatButton.cls"
$macroRowClassText = Read-VbaSource "class_modules\clsMacroListRow.cls"

$frmQuickText = Read-VbaSource "userforms\frmQuickToolbar.frm"
$frmMacroText = Read-VbaSource "userforms\frmMacroList.frm"
$frmBatchText = Read-VbaSource "userforms\frmBatchMode.frm"
$frmReviewerText = Read-VbaSource "userforms\frmReviewerFilter.frm"
$frmCaptionText = Read-VbaSource "userforms\frmCaptionReference.frm"
$frmAboutText = Read-VbaSource "userforms\frmAbout.frm"

# --- Core_Window DWM & 光标扩展 ---
Assert-Contains $windowText "DwmSetWindowAttribute" "Core_Window should declare DwmSetWindowAttribute."
Assert-Contains $windowText "DWMWA_WINDOW_CORNER_PREFERENCE" "Core_Window should configure Win11 window corners."
Assert-Contains $windowText "DWMWA_CAPTION_COLOR" "Core_Window should configure Win11 caption color."
Assert-Contains $windowText "Optional ByVal applyCaptionColor As Boolean = True" "Core_Window should allow skipping caption color for WS_EX_TOOLWINDOW."
Assert-Contains $windowText "Public Function StyleUserFormWindowChrome(" "Core_Window should export StyleUserFormWindowChrome."
Assert-Contains $windowText "Public Sub SetHandCursor()" "Core_Window should export SetHandCursor."

# --- UI_Theme 双主题探测与配色板 ---
foreach ($required in @(
    "RegOpenKeyEx", "RegQueryValueEx", "RegCloseKey",
    "UI Theme", "AppsUseLightTheme",
    "Public Function IsOfficeDarkTheme() As Boolean",
    "Public Sub SetThemeOverrideForTest(",
    "RGB(245, 247, 250)", "RGB(24, 90, 189)",
    "RGB(32, 34, 37)", "RGB(79, 156, 249)",
    "Public Sub ApplyFormBaseTheme(",
    "Public Sub ApplyFormWindowChrome(",
    "Public Function CreateFlatButton(",
    "Public Function RegisterHoverButton(",
    "Public Sub RefreshHoverButtons(",
    "Public Sub ResetHoverButtons(",
    "Public Sub ClearHoverButtons("
)) {
    Assert-Contains $themeText $required "UI_Theme missing expected element: $required"
}

# --- 交互类检查 ---
foreach ($required in @(
    "Public WithEvents Lbl As MSForms.Label",
    "Public WithEvents Btn As MSForms.CommandButton",
    "Public Sub ApplyVisualState()",
    "Public Sub ResetHover()",
    "SetHandCursor"
)) {
    Assert-Contains $hoverClassText $required "clsHoverButton missing expected element: $required"
}

foreach ($required in @(
    "Public WithEvents m_Hit As MSForms.Label",
    "Public Sub Init(",
    "Public Sub MoveTo(",
    "Public Sub ApplyVisualState()",
    "Public Sub Dispose()",
    "OnFlatButtonClick"
)) {
    Assert-Contains $flatBtnClassText $required "clsFlatButton missing expected element: $required"
}

foreach ($required in @(
    "Public WithEvents m_Hit As MSForms.Label",
    "Public Sub Init(",
    "Public Sub SetSelected(",
    "Public Sub Dispose()",
    "OnMacroRowClick"
)) {
    Assert-Contains $macroRowClassText $required "clsMacroListRow missing expected element: $required"
}

# --- 全部 6 个 UserForm 统一接线主题、扁平按钮与 Win11 外观 ---
$forms = @{
    "frmQuickToolbar"     = $frmQuickText
    "frmMacroList"        = $frmMacroText
    "frmBatchMode"        = $frmBatchText
    "frmReviewerFilter"   = $frmReviewerText
    "frmCaptionReference" = $frmCaptionText
    "frmAbout"            = $frmAboutText
}

foreach ($entry in $forms.GetEnumerator()) {
    Assert-Contains $entry.Value "ApplyFormBaseTheme Me" "$($entry.Key) should call ApplyFormBaseTheme Me."
    Assert-True ($entry.Value.IndexOf("ApplyFormWindowChrome Me.Caption", [StringComparison]::OrdinalIgnoreCase) -ge 0) "$($entry.Key) should call ApplyFormWindowChrome Me.Caption."
    Assert-True ($entry.Value.Contains("RegisterHoverButton") -or $entry.Value.Contains("CreateFlatButton")) "$($entry.Key) should register hover/flat buttons."
    Assert-Contains $entry.Value "UserForm_Activate()" "$($entry.Key) should refresh chrome/theme on activation."
    Assert-Contains $entry.Value "ClearHoverButtons" "$($entry.Key) should release hover/flat button references on close."
}

Assert-Contains $frmQuickText "ApplyFormWindowChrome Me.Caption, False" "frmQuickToolbar should avoid DWM caption color on WS_EX_TOOLWINDOW."
Assert-Contains $frmMacroText "clsMacroListRow" "frmMacroList should use clsMacroListRow custom padded rows."
Assert-Contains $frmBatchText "lblAccentBar" "frmBatchMode should render active left accent bars."

# --- frmBatchMode 资源与 Core_UI 接入 ---
$batchFrx = Join-Path $repoRoot "userforms\frmBatchMode.frx"
Assert-True (Test-Path -LiteralPath $batchFrx -PathType Leaf) "frmBatchMode.frx should exist."
Assert-True ((Get-Item -LiteralPath $batchFrx).Length -gt 1000) "frmBatchMode.frx should be non-empty."
Assert-Contains $coreUiText 'VBA.UserForms.Add("frmBatchMode")' "Core_UI.ChooseBatchMode should invoke frmBatchMode."
Assert-Contains $coreUiText "ChooseBatchModeLegacy" "Core_UI should retain InputBox fallback for standalone tests."

# --- 悬浮窗按钮计数保护与 Sync 脚本一致性 ---
Assert-Contains $qtModuleText 'If Left$(ctrl.Name, 3) = "btn" Then' "GetQuickToolbarButtonCount should count only btn* controls."
Assert-Contains $qtModuleText "frmQuickToolbar.ApplyTheme True" "ShowQuickToolbar should apply current Word theme."
Assert-Contains $qtModuleText "frmQuickToolbar.ApplyTheme False" "RaiseQuickToolbarIfVisible should sync Word theme changes."

$syncText = [System.IO.File]::ReadAllText((Join-Path $repoRoot "scripts\Sync-QuickToolbarForm.ps1"), [System.Text.Encoding]::UTF8)
$buttonData = Import-RAToolsDataFile (Join-Path $repoRoot "scripts\QuickToolbarButtons.psd1")
$layoutSpec = $buttonData.Layout
$toolbarItems = $buttonData.Buttons

$tplMatch = [regex]::Match($syncText, "(?s)\`$template = @'\r?\n(.*?)\r?\n'@")
Assert-True $tplMatch.Success "Sync-QuickToolbarForm.ps1 should contain VBA template."

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
$gapLines = foreach ($gapRow in $layoutSpec.GroupGapRows) {
    "    If rowIndex >= $gapRow Then GetButtonTop = GetButtonTop + BUTTON_GROUP_GAP"
}
$expectedQuickBody = $tplMatch.Groups[1].Value.
    Replace("__TOOLBAR_WIDTH__", [string]$layoutSpec.ToolbarWidth).
    Replace("__TOOLBAR_HEIGHT__", [string]$layoutSpec.ToolbarHeight).
    Replace("__BUTTON_TOP__", [string]$layoutSpec.ButtonTop).
    Replace("__BUTTON_WIDTH__", [string]$layoutSpec.ButtonWidth).
    Replace("__BUTTON_HEIGHT__", [string]$layoutSpec.ButtonHeight).
    Replace("__BUTTON_ROW_STEP__", [string]$layoutSpec.RowStep).
    Replace("__BUTTON_COLUMN_GAP__", [string]$layoutSpec.ColumnGap).
    Replace("__BUTTON_GROUP_GAP__", [string]$layoutSpec.GroupGap).
    Replace("__CONFIGURE_LINES__", ($configureLines -join "`r`n")).
    Replace("__GROUP_GAP_LINES__", ($gapLines -join "`r`n")).
    Replace("__GROUP_GAP_LIST__", ($layoutSpec.GroupGapRows -join ", ")).
    Replace("__HANDLERS__", ($handlerLines -join "`r`n"))

$actualQuickBody = $frmQuickText.Substring($frmQuickText.IndexOf("Option Explicit", [StringComparison]::Ordinal)).TrimEnd("`r", "`n")
Assert-True (($expectedQuickBody -replace "`r?`n", "`r`n").TrimEnd("`r", "`n") -eq $actualQuickBody) `
    "userforms\frmQuickToolbar.frm code body should stay synchronized with scripts\Sync-QuickToolbarForm.ps1."

Write-Output "PASS UITheme.Tests ($assertions assertions)"
