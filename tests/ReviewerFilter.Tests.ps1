$ErrorActionPreference = 'Stop'

# 「按审阅者筛选」功能测试：
# 1) 静态契约：模块/窗体/功能区/宏注册表接线齐全；
# 2) Word COM 真实行为：多作者修订与批注的作者收集、按作者读写显示状态。

Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $repoRoot "modules\Macro_ReviewerFilter.bas"
$formPath = Join-Path $repoRoot "userforms\frmReviewerFilter.frm"
$ribbonPath = Join-Path $repoRoot "dotm\customUI\customUI14.xml"
$registryPath = Join-Path $repoRoot "modules\UI_MacroRegistry.bas"

function Assert-True {
    param([bool]$Condition, [string]$Message)

    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Read-VbaSource {
    param([string]$Path)

    $vbaEncoding = [System.Text.Encoding]::GetEncoding(936)
    return $vbaEncoding.GetString([System.IO.File]::ReadAllBytes($Path))
}

Write-Host "Running static contract checks"

Assert-True (Test-Path -LiteralPath $modulePath) "Missing module: $modulePath"
Assert-True (Test-Path -LiteralPath $formPath) "Missing form: $formPath"

$moduleSource = Read-VbaSource $modulePath
$formSource = Read-VbaSource $formPath

foreach ($required in @('Public Sub ShowReviewerFilter(', 'Public Sub Wrapper_ShowReviewerFilter(',
        'Public Function CollectReviewerNames(', 'Public Function ApplyReviewerFilter(',
        'Public Function IsReviewerVisible(', 'Public Sub RefreshReviewerFilterIfVisible(')) {
    Assert-True ($moduleSource.Contains($required)) "Module is missing $required"
}

foreach ($required in @('lstAuthors_Click', 'lstAuthors_DblClick', 'chkAll_MouseUp',
        'Public Sub RefreshReviewerList(', 'Public Sub ToggleAllReviewers(',
        'Public Sub ShowOnlyAuthorAt(', 'Public Sub ToggleAuthorAt(',
        'ApplyReviewerFilter(')) {
    Assert-True ($formSource.Contains($required)) "Form is missing $required"
}

$ribbonText = Get-Content -LiteralPath $ribbonPath -Raw
Assert-True ($ribbonText.Contains('onAction="ShowReviewerFilter"')) 'Ribbon button is not wired'
Assert-True ($ribbonText.Contains('id="btnReviewerFilter"')) 'Ribbon button id missing'

$registrySource = Read-VbaSource $registryPath
Assert-True ($registrySource.Contains('Wrapper_ShowReviewerFilter')) 'Macro registry entry missing'

Write-Host "PASS static contract checks"

$word = $null
$doc = $null

try {
    $word = Start-RAToolsWordSession
    $doc = $word.Documents.Add()

    [void]$doc.VBProject.VBComponents.Import($modulePath)
    $form = $doc.VBProject.VBComponents.Import($formPath)
    Assert-True ([int]$form.Designer.Controls.Count -ge 4) "Form controls were not imported"

    $originalUserName = $word.UserName

    # 三位作者：Alice 正文修订、Bob 页眉修订、Carol 批注
    $doc.TrackRevisions = $true
    $word.UserName = 'Alice'
    $doc.Content.Text = 'main text '
    $word.UserName = 'Bob'
    $doc.Sections(1).Headers(1).Range.Text = 'header '
    $word.UserName = 'Carol'
    $doc.Comments.Add($doc.Content, 'carol comment') | Out-Null

    $helper = $doc.VBProject.VBComponents.Add(1)
    $helper.Name = 'ReviewerFilterProbe'
    $helper.CodeModule.AddFromString(@'
Public Sub RAToolsReviewerFilterProbe()
    Dim names() As String
    Dim counts() As Long
    Dim total As Long
    Dim flags() As Boolean
    Dim i As Long

    On Error Resume Next
    ThisDocument.Variables("ReviewerNames").Delete
    ThisDocument.Variables("ReviewerCounts").Delete
    ThisDocument.Variables("ReviewerApplyResult").Delete
    ThisDocument.Variables("ReviewerVisibleAfter").Delete
    ThisDocument.Variables("ReviewerError").Delete
    On Error GoTo 0

    names = CollectReviewerNames(ThisDocument, counts, total)
    ThisDocument.Variables.Add "ReviewerNames", Join(names, "|")
    ThisDocument.Variables.Add "ReviewerCounts", CStr(total)

    If total > 0 Then
        ReDim flags(0 To total - 1)
    Else
        ReDim flags(0 To 0)
    End If
    For i = 0 To total - 1
        flags(i) = False
    Next i
    ' 只显示第一位作者
    If total > 0 Then flags(0) = True

    ThisDocument.Variables.Add "ReviewerApplyResult", CStr(ApplyReviewerFilter(ThisDocument, names, flags, total))

    Dim state As String
    For i = 0 To total - 1
        state = state & names(i) & "=" & IsReviewerVisible(ThisDocument, names(i)) & ";"
    Next i
    ThisDocument.Variables.Add "ReviewerVisibleAfter", state
    ThisDocument.Variables.Add "ReviewerError", CStr(Err.Number)
End Sub
'@)

    [void]$word.Run('RAToolsReviewerFilterProbe')

    $names = [string]$doc.Variables.Item("ReviewerNames").Value
    $total = [int]$doc.Variables.Item("ReviewerCounts").Value
    $applyResult = [string]$doc.Variables.Item("ReviewerApplyResult").Value
    $visibleAfter = [string]$doc.Variables.Item("ReviewerVisibleAfter").Value

    Write-Host "Collected authors: $names"
    Write-Host "Visible after filter: $visibleAfter"

    Assert-True ($total -eq 3) "Expected 3 reviewers, got $total"
    foreach ($author in @('Alice', 'Bob', 'Carol')) {
        Assert-True ($names.Contains($author)) "Collector missed author $author"
    }
    Assert-True ($applyResult -eq 'True') "ApplyReviewerFilter reported failure"
    Assert-True ($visibleAfter.Contains('Alice=True')) 'First reviewer should stay visible'
    Assert-True ($visibleAfter.Contains('Bob=False')) 'Bob should be hidden'
    Assert-True ($visibleAfter.Contains('Carol=False')) 'Carol should be hidden'

    # 恢复全部显示，确认可逆
    $doc.ActiveWindow.View.Reviewers.Item('Bob').Visible = $true
    $doc.ActiveWindow.View.Reviewers.Item('Carol').Visible = $true
    Assert-True ([bool]$doc.ActiveWindow.View.Reviewers.Item('Bob').Visible) 'Bob should be visible again'

    $word.UserName = $originalUserName
    $doc.Close(0)
    $doc = $null

    Write-Host "PASS ReviewerFilter.Tests"
    exit 0
}
catch {
    Write-Host ("FAIL ReviewerFilter.Tests: " + $_.Exception.Message) -ForegroundColor Red
    if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
    exit 1
}
finally {
    Stop-RAToolsWordSession -Word $word
}
