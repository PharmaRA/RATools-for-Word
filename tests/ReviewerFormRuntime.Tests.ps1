$ErrorActionPreference = 'Stop'

# 窗体运行时冒烟测试：加载 frmReviewerFilter，验证
#   1) 「全部审阅者」与各行勾选状态始终一致（点任意一行都会同步）；
#   2) 单击切换、只看某人、全选/全不选都真实写回文档；
#   3) 同一文档重复刷新命中缓存，切窗口不再重复枚举修订。

Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $repoRoot "modules\Macro_ReviewerFilter.bas"
$itemPath = Join-Path $repoRoot "class_modules\clsReviewerItem.cls"
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
    [void]$doc.VBProject.VBComponents.Import($itemPath)
    [void]$doc.VBProject.VBComponents.Import($formPath)

    $originalUserName = $word.UserName

    $doc.TrackRevisions = $true
    $word.UserName = 'Alice'
    $doc.Content.Text = 'main text '
    $word.UserName = 'Bob'
    $doc.Content.InsertAfter('more text ')
    $word.UserName = 'Carol'
    $doc.Content.InsertAfter('last text ')
    $doc.TrackRevisions = $false

    $helper = $doc.VBProject.VBComponents.Add(1)
    $helper.Name = 'ReviewerFormProbe'
    $helper.CodeModule.AddFromString(@'
Private Function CheckOf(ByVal authorIndex As Long) As Boolean
    Dim c As Control
    Dim n As Long

    n = 0
    For Each c In frmReviewerFilter.frmList.Controls
        If n = authorIndex + 1 Then
            CheckOf = (c.Value = True)
            Exit Function
        End If
        n = n + 1
    Next c
End Function

Public Sub RAToolsReviewerFormProbe()
    Dim state As String
    Dim buildsBefore As Long

    On Error Resume Next
    ThisDocument.Variables("FormProbeResult").Delete
    ThisDocument.Variables("FormProbeError").Delete
    On Error GoTo 0

    On Error GoTo ErrH

    Load frmReviewerFilter
    frmReviewerFilter.RefreshReviewerList False

    state = state & "rows=" & frmReviewerFilter.frmList.Controls.Count & ";"
    state = state & "allChecked=" & frmReviewerFilter.frmList.Controls(0).Value & ";"
    state = state & "onlyEnabledAtStart=" & frmReviewerFilter.btnOnlyThis.Enabled & ";"

    ' 单击 Bob（真实路径：翻转控件值 -> 控件触发 Click -> clsReviewerItem 转发）：
    ' 只隐藏 Bob，且「全部审阅者」自动取消勾选
    Dim chk As Object
    Set chk = frmReviewerFilter.frmList.Controls("chkRow2")
    chk.Value = Not (chk.Value = True)
    state = state & "onlyEnabledAfterClick=" & frmReviewerFilter.btnOnlyThis.Enabled & ";"
    state = state & "bobDoc=" & IsReviewerVisible(ThisDocument, "Bob") & ";"
    state = state & "bobChk=" & CheckOf(1) & ";"
    state = state & "aliceDoc=" & IsReviewerVisible(ThisDocument, "Alice") & ";"
    state = state & "allAfterOne=" & frmReviewerFilter.frmList.Controls(0).Value & ";"

    ' 再次单击 Bob 恢复
    Set chk = frmReviewerFilter.frmList.Controls("chkRow2")
    chk.Value = Not (chk.Value = True)
    state = state & "bobBack=" & IsReviewerVisible(ThisDocument, "Bob") & ";"
    state = state & "allBack=" & frmReviewerFilter.frmList.Controls(0).Value & ";"

    ' 「全部审阅者」全关再全开
    Set chk = frmReviewerFilter.frmList.Controls("chkRow0")
    chk.Value = False
    state = state & "allOffAlice=" & IsReviewerVisible(ThisDocument, "Alice") & ";"
    state = state & "allOffCarol=" & IsReviewerVisible(ThisDocument, "Carol") & ";"
    state = state & "allOffChk2=" & CheckOf(2) & ";"

    Set chk = frmReviewerFilter.frmList.Controls("chkRow0")
    chk.Value = True
    state = state & "allOnAlice=" & IsReviewerVisible(ThisDocument, "Alice") & ";"
    state = state & "allOnCarol=" & IsReviewerVisible(ThisDocument, "Carol") & ";"
    state = state & "allOnChk=" & frmReviewerFilter.frmList.Controls(0).Value & ";"

    ' 只看第三人：先单击选中他，再点「只看此人」
    frmReviewerFilter.ToggleAuthorAt 2
    frmReviewerFilter.ToggleAuthorAt 2
    frmReviewerFilter.ShowOnlyLastClicked
    state = state & "onlyAlice=" & IsReviewerVisible(ThisDocument, "Alice") & ";"
    state = state & "onlyBob=" & IsReviewerVisible(ThisDocument, "Bob") & ";"
    state = state & "onlyCarol=" & IsReviewerVisible(ThisDocument, "Carol") & ";"
    state = state & "onlyChk0=" & CheckOf(0) & ";"
    state = state & "onlyChk2=" & CheckOf(2) & ";"
    state = state & "onlyAll=" & frmReviewerFilter.frmList.Controls(0).Value & ";"

    ' 缓存：同一文档重复刷新不应再次枚举修订
    buildsBefore = GetSnapshotBuildCountForTest()
    frmReviewerFilter.RefreshReviewerList False
    frmReviewerFilter.RefreshReviewerList False
    state = state & "buildsDelta=" & (GetSnapshotBuildCountForTest() - buildsBefore) & ";"

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

    Assert-True ($state.Contains('rows=4')) 'Panel must show the all-reviewers row plus three authors'
    Assert-True ($state.Contains('onlyEnabledAtStart=False')) 'Only-this button must start disabled'
    Assert-True ($state.Contains('onlyEnabledAfterClick=True')) 'Only-this button must enable after a click'
    Assert-True ($state.Contains('allChecked=True')) 'All-reviewers must start checked'

    Assert-True ($state.Contains('bobDoc=False')) 'Clicking a reviewer must hide it'
    Assert-True ($state.Contains('bobChk=False')) 'The clicked checkbox must reflect the new state'
    Assert-True ($state.Contains('aliceDoc=True')) 'Other reviewers must stay visible'
    Assert-True ($state.Contains('allAfterOne=False')) 'All-reviewers must uncheck on partial selection'

    Assert-True ($state.Contains('bobBack=True')) 'Second click must show the reviewer again'
    Assert-True ($state.Contains('allBack=True')) 'All-reviewers must recheck when all are shown'

    Assert-True ($state.Contains('allOffAlice=False')) 'All-off must hide Alice'
    Assert-True ($state.Contains('allOffCarol=False')) 'All-off must hide Carol'
    Assert-True ($state.Contains('allOffChk2=False')) 'All-off must clear every row checkbox'
    Assert-True ($state.Contains('allOnAlice=True')) 'All-on must show Alice'
    Assert-True ($state.Contains('allOnCarol=True')) 'All-on must show Carol'
    Assert-True ($state.Contains('allOnChk=True')) 'All-on must check every row checkbox'

    Assert-True ($state.Contains('onlyAlice=False')) 'Show-only must hide Alice'
    Assert-True ($state.Contains('onlyBob=False')) 'Show-only must hide Bob'
    Assert-True ($state.Contains('onlyCarol=True')) 'Show-only must show Carol'
    Assert-True ($state.Contains('onlyChk0=False')) 'Show-only must leave other rows unchecked'
    Assert-True ($state.Contains('onlyChk2=True')) 'Show-only must check the chosen row'
    Assert-True ($state.Contains('onlyAll=False')) 'Show-only must leave the all-reviewers row unchecked'

    Assert-True ($state.Contains('buildsDelta=0')) 'Repeated refresh for the same document must hit the cache'

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
