$ErrorActionPreference = 'Stop'

# 窗体运行时冒烟测试：加载 frmReviewerFilter、刷新列表、全选/取消全选，
# 断言底层 Reviewers 显示状态随之变化。不显示窗体（保持自动化无交互）。

Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $repoRoot "modules\Macro_ReviewerFilter.bas"
$formPath = Join-Path $repoRoot "userforms\frmReviewerFilter.frm"

function Assert-True {
    param([bool]$Condition, [string]$Message)

    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

$word = $null
$doc = $null

try {
    $word = Start-RAToolsWordSession
    $doc = $word.Documents.Add()

    [void]$doc.VBProject.VBComponents.Import($modulePath)
    [void]$doc.VBProject.VBComponents.Import($formPath)

    $originalUserName = $word.UserName

    $doc.TrackRevisions = $true
    $word.UserName = 'Alice'
    $doc.Content.Text = 'main text '
    $word.UserName = 'Bob'
    $doc.Content.InsertAfter('more text ')

    $helper = $doc.VBProject.VBComponents.Add(1)
    $helper.Name = 'ReviewerFormProbe'
    $helper.CodeModule.AddFromString(@'
Public Sub RAToolsReviewerFormProbe()
    Dim state As String

    On Error Resume Next
    ThisDocument.Variables("FormProbeResult").Delete
    ThisDocument.Variables("FormProbeError").Delete
    On Error GoTo 0

    On Error GoTo ErrH

    Load frmReviewerFilter
    frmReviewerFilter.RefreshReviewerList
    state = state & "items=" & frmReviewerFilter.lstAuthors.ListCount & ";"
    state = state & "allChecked=" & frmReviewerFilter.chkAll.Value & ";"

    ' 取消「全部审阅者」：两位作者都应隐藏
    frmReviewerFilter.chkAll.Value = False
    frmReviewerFilter.ToggleAllReviewers
    state = state & "alice=" & IsReviewerVisible(ThisDocument, "Alice") & ";"
    state = state & "bob=" & IsReviewerVisible(ThisDocument, "Bob") & ";"

    ' 勾上「全部审阅者」：两位作者都应恢复显示
    frmReviewerFilter.chkAll.Value = True
    frmReviewerFilter.ToggleAllReviewers
    state = state & "aliceBack=" & IsReviewerVisible(ThisDocument, "Alice") & ";"
    state = state & "bobBack=" & IsReviewerVisible(ThisDocument, "Bob") & ";"

    ' 只看某人：双击语义
    frmReviewerFilter.ShowOnlyAuthorAt 1
    state = state & "onlyAlice=" & IsReviewerVisible(ThisDocument, "Alice") & ";"
    state = state & "onlyBob=" & IsReviewerVisible(ThisDocument, "Bob") & ";"

    Unload frmReviewerFilter
    ThisDocument.Variables.Add "FormProbeResult", state
    Exit Sub

ErrH:
    ThisDocument.Variables.Add "FormProbeError", CStr(Err.Number) & " " & Err.Description
End Sub
'@)

    [void]$word.Run('RAToolsReviewerFormProbe')

    $errorText = ''
    try { $errorText = [string]$doc.Variables.Item("FormProbeError").Value } catch { }
    Assert-True ([string]::IsNullOrEmpty($errorText)) "Form probe raised: $errorText"

    $state = [string]$doc.Variables.Item("FormProbeResult").Value
    Write-Host "Form state: $state"

    Assert-True ($state.Contains('items=2')) 'List did not show both reviewers'
    Assert-True ($state.Contains('allChecked=True')) 'All-reviewers checkbox should start checked'
    Assert-True ($state.Contains('alice=False')) 'Unchecking all should hide Alice'
    Assert-True ($state.Contains('bob=False')) 'Unchecking all should hide Bob'
    Assert-True ($state.Contains('aliceBack=True')) 'Rechecking all should show Alice'
    Assert-True ($state.Contains('bobBack=True')) 'Rechecking all should show Bob'
    Assert-True ($state.Contains('onlyAlice=False')) 'Show-only should hide Alice'
    Assert-True ($state.Contains('onlyBob=True')) 'Show-only should show Bob'

    $word.UserName = $originalUserName
    $doc.Close(0)
    $doc = $null

    Write-Host "PASS ReviewerFormRuntime.Tests"
    exit 0
}
catch {
    Write-Host ("FAIL ReviewerFormRuntime.Tests: " + $_.Exception.Message) -ForegroundColor Red
    if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
    exit 1
}
finally {
    Stop-RAToolsWordSession -Word $word
}
