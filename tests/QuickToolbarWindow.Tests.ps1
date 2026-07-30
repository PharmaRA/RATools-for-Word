$ErrorActionPreference = "Stop"

# 悬浮窗窗口层契约（无 COM）：断言"脱离文档窗口 + 置顶 + 交还焦点"三件事在源码
# 里始终成对出现。这些行为只在真实 Word 里能肉眼验证，而 CI 没有 Word，因此这里
# 用源码级断言守住调用关系与顺序，防止再次退化成"点一次按钮悬浮窗就沉到后台"。
# 同时校验三个文件仍是合法 GBK（见 docs/vba-source-encoding.md）。

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
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

function Assert-NotContains {
    param(
        [string]$Text,
        [string]$Needle,
        [string]$Message
    )

    Assert-True (-not $Text.Contains($Needle)) $Message
}

function Read-VbaSource {
    param([string]$Path)

    $gbk = [System.Text.Encoding]::GetEncoding(936)
    return $gbk.GetString([System.IO.File]::ReadAllBytes($Path))
}

function Get-ProcedureBody {
    param(
        [string]$Text,
        [string]$Signature,
        [string]$Terminator
    )

    $pattern = "(?s)" + [regex]::Escape($Signature) + ".*?" + [regex]::Escape($Terminator)
    $match = [regex]::Match($Text, $pattern)
    if (-not $match.Success) {
        throw "ASSERTION FAILED: could not find procedure <$Signature>"
    }
    return $match.Value
}

function Assert-CallOrder {
    param(
        [string]$Body,
        [string]$First,
        [string]$Then,
        [string]$Message
    )

    $firstIndex = $Body.IndexOf($First, [StringComparison]::Ordinal)
    $thenIndex = $Body.IndexOf($Then, [StringComparison]::Ordinal)
    Assert-True ($firstIndex -ge 0 -and $thenIndex -gt $firstIndex) $Message
}

Write-Output "Running QuickToolbar window-layer checks"

$windowModulePath = Join-Path $repoRoot "modules\Mod_Core_Window.bas"
$toolbarModulePath = Join-Path $repoRoot "modules\Mod_QuickToolbar.bas"
$appEventsPath = Join-Path $repoRoot "class_modules\clsAppEvents.cls"

foreach ($requiredPath in @($windowModulePath, $toolbarModulePath, $appEventsPath)) {
    Assert-True (Test-Path -LiteralPath $requiredPath -PathType Leaf) "Required source file should exist: $requiredPath"
}

# --- GBK 契约：三个文件必须仍是无损的 936 编码 ---
$gbk = [System.Text.Encoding]::GetEncoding(936)
foreach ($sourcePath in @($windowModulePath, $toolbarModulePath, $appEventsPath)) {
    $bytes = [System.IO.File]::ReadAllBytes($sourcePath)
    $decoded = $gbk.GetString($bytes)
    $roundTrip = $gbk.GetBytes($decoded)
    $name = Split-Path -Leaf $sourcePath
    Assert-True ($roundTrip.Length -eq $bytes.Length) "$name should still be lossless GBK(936); see docs/vba-source-encoding.md."
    Assert-True (-not $decoded.Contains([char]0xFFFD)) "$name should not contain replacement characters (UTF-8 written over GBK)."
    Assert-NotContains $decoded "`n`n`n" "$name should not gain stray blank-line runs from a bad edit."
}

$windowText = Read-VbaSource -Path $windowModulePath
$toolbarText = Read-VbaSource -Path $toolbarModulePath
$appEventsText = Read-VbaSource -Path $appEventsPath

# --- Mod_Core_Window：置顶层原语 ---
Assert-Contains $windowText 'Private Declare PtrSafe Function SetWindowPos Lib "user32"' "Window layer should declare SetWindowPos."
Assert-Contains $windowText "Private Const HWND_TOPMOST As LongPtr = -1" "Window layer should define HWND_TOPMOST."
Assert-Contains $windowText "Private Const WS_EX_TOPMOST As Long = &H8" "Window layer should define WS_EX_TOPMOST for its state check."
Assert-Contains $windowText "Public Function KeepUserFormWindowOnTop(ByVal windowCaption As String) As Boolean" "Window layer should expose KeepUserFormWindowOnTop."
Assert-Contains $windowText "Public Function IsUserFormWindowOnTop(ByVal windowCaption As String) As Boolean" "Window layer should expose IsUserFormWindowOnTop."
Assert-NotContains $windowText "RaiseUserFormWindow" "HWND_TOP based raising was replaced by the topmost band; no stale helper should remain."
Assert-NotContains $windowText "HWND_TOP As LongPtr = 0" "Plain HWND_TOP cannot keep an ownerless form above Word; it should be gone."

$keepOnTopBody = Get-ProcedureBody -Text $windowText `
    -Signature "Public Function KeepUserFormWindowOnTop(ByVal windowCaption As String) As Boolean" `
    -Terminator "End Function"
Assert-Contains $keepOnTopBody "SetWindowPos windowHandle, HWND_TOPMOST" "KeepUserFormWindowOnTop should insert the form into the topmost band."
Assert-Contains $keepOnTopBody "SWP_NOACTIVATE" "Raising the toolbar must never steal the editing focus."
Assert-Contains $keepOnTopBody "SWP_NOMOVE Or SWP_NOSIZE" "Raising the toolbar must not move or resize it."
Assert-Contains $keepOnTopBody "KeepUserFormWindowOnTop = IsUserFormWindowOnTop(windowCaption)" "KeepUserFormWindowOnTop should report the resulting state."

$isOnTopBody = Get-ProcedureBody -Text $windowText `
    -Signature "Public Function IsUserFormWindowOnTop(ByVal windowCaption As String) As Boolean" `
    -Terminator "End Function"
Assert-Contains $isOnTopBody "GetWindowLongCompat(windowHandle, GWL_EXSTYLE) And WS_EX_TOPMOST" "IsUserFormWindowOnTop should read the live WS_EX_TOPMOST bit."

# 脱离 owner 与置顶必须成对：owner 为 0 的窗体不再被 Word 自动压在文档窗口之上
$detachBody = Get-ProcedureBody -Text $windowText `
    -Signature "Public Function DetachUserFormWindowFromOwner(ByVal windowCaption As String) As Boolean" `
    -Terminator "End Function"
Assert-Contains $detachBody "SetWindowLongCompat windowHandle, GWLP_HWNDPARENT, 0" "Detach should clear the owner window."
Assert-Contains $detachBody "WS_EX_TOOLWINDOW" "Detach should keep the toolbar out of Alt+Tab and the taskbar."
Assert-CallOrder -Body $detachBody -First "ShowWindow windowHandle, SW_SHOWNA" -Then "KeepUserFormWindowOnTop windowCaption" `
    -Message "Detach should pin the form on top after re-showing it, otherwise Word buries an ownerless toolwindow."

# --- Mod_QuickToolbar：显示、动作、焦点归还 ---
$showBody = Get-ProcedureBody -Text $toolbarText -Signature "Public Sub ShowQuickToolbar()" -Terminator "End Sub"
Assert-CallOrder -Body $showBody -First "frmQuickToolbar.Show vbModeless" -Then "DetachQuickToolbarFromDocumentWindow" `
    -Message "ShowQuickToolbar should detach only after the window handle exists."
Assert-CallOrder -Body $showBody -First "DetachQuickToolbarFromDocumentWindow" -Then "ReturnFocusToDocumentWindow" `
    -Message "ShowQuickToolbar should hand the focus back to the document last."

$actionBody = Get-ProcedureBody -Text $toolbarText -Signature "Public Sub RunQuickToolbarAction(ByVal actionKey As String)" -Terminator "End Sub"
Assert-CallOrder -Body $actionBody -First "DispatchQuickToolbarAction actionKey" -Then "ReturnFocusToDocumentWindow" `
    -Message "Every toolbar action should return the focus to the document when it finishes."

$raiseBody = Get-ProcedureBody -Text $toolbarText -Signature "Public Sub RaiseQuickToolbarIfVisible()" -Terminator "End Sub"
Assert-Contains $raiseBody "If Not IsUserFormWindowVisible(QUICK_TOOLBAR_WINDOW_CAPTION) Then Exit Sub" "Raising should no-op while the toolbar is hidden."
Assert-Contains $raiseBody "KeepUserFormWindowOnTop QUICK_TOOLBAR_WINDOW_CAPTION" "Raising should re-assert the topmost band."
Assert-NotContains $raiseBody "frmQuickToolbar.Visible" "Raising must not touch the default instance, or it would load the form implicitly."

$focusBody = Get-ProcedureBody -Text $toolbarText -Signature "Public Sub ReturnFocusToDocumentWindow()" -Terminator "End Sub"
Assert-Contains $focusBody "If Documents.Count = 0 Then Exit Sub" "Returning focus should be silent with no document open."
Assert-CallOrder -Body $focusBody -First "Application.Activate" -Then "RaiseQuickToolbarIfVisible" `
    -Message "Activating Word raises its document window, so the toolbar must be re-asserted afterwards; without this the toolbar sinks behind Word after one click."

Assert-Contains $toolbarText "Public Sub UnloadQuickToolbarOnQuit()" "Word quit should still unload the ownerless toolbar."

# --- clsAppEvents：文档生命周期钩子 ---
$documentChangeBody = Get-ProcedureBody -Text $appEventsText -Signature "Private Sub App_DocumentChange()" -Terminator "End Sub"
Assert-Contains $documentChangeBody "Mod_QuickToolbar.RaiseQuickToolbarIfVisible" "Switching documents should re-assert the toolbar."
$windowActivateBody = Get-ProcedureBody -Text $appEventsText -Signature "Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)" -Terminator "End Sub"
Assert-Contains $windowActivateBody "Mod_QuickToolbar.RaiseQuickToolbarIfVisible" "Activating a document window should re-assert the toolbar."
$quitBody = Get-ProcedureBody -Text $appEventsText -Signature "Private Sub App_Quit()" -Terminator "End Sub"
Assert-Contains $quitBody "Mod_QuickToolbar.UnloadQuickToolbarOnQuit" "Word quit should save the toolbar position."

# --- 跨文件引用完整性：悬浮窗调用的窗口层过程必须真实存在 ---
# 重命名窗口层过程而漏改调用方，只有在 Word 里编译时才会报错，这里提前拦住。
$exported = @{}
foreach ($match in [regex]::Matches($windowText, "(?m)^Public (?:Function|Sub) (\w+)")) {
    $exported[$match.Groups[1].Value] = $true
}
Assert-True ($exported.Count -ge 5) "Window layer should export its helper set; found $($exported.Count)."

$callers = @{ "Mod_QuickToolbar.bas" = $toolbarText }
foreach ($caller in $callers.GetEnumerator()) {
    foreach ($match in [regex]::Matches($caller.Value, "\b(\w*UserFormWindow\w*)\b")) {
        $name = $match.Groups[1].Value
        Assert-True ($exported.ContainsKey($name)) "$($caller.Key) calls $name, which Mod_Core_Window does not export."
    }
}

Write-Output "PASS QuickToolbarWindow.Tests ($assertions assertions)"
