#requires -version 5.1
[CmdletBinding()]
param([ValidateSet('Writing','Final')][string[]]$Kind = @('Writing'))
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'RATools.Build.psm1') -Force
$layout = Import-PowerShellDataFile (Join-Path $PSScriptRoot 'WorkflowToolbarLayout.psd1')
$gbk = [Text.Encoding]::GetEncoding(936)
function VbaQuote([string]$Text) { return '"' + $Text.Replace('"', '""') + '"' }
foreach ($toolbarKind in $Kind) {
    $definition = $layout[$toolbarKind]
    $formName = 'frm' + $toolbarKind + 'Toolbar'
    $backupRoot = Join-Path $repoRoot ('backups/workflow-form-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($backupRoot)
    foreach ($extension in @('frm','frx')) {
        $existing = Join-Path $repoRoot "userforms/$formName.$extension"
        if (Test-Path -LiteralPath $existing) { Copy-Item -LiteralPath $existing -Destination $backupRoot }
    }
    $word = $null
    $document = $null
    $exportRoot = Join-Path ([IO.Path]::GetTempPath()) ('RATools_WorkflowForm_' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($exportRoot)
    try {
        $word = Start-RAToolsWordSession
        $document = $word.Documents.Add()
        $component = $document.VBProject.VBComponents.Add(3)
        $component.Name = $formName
        $component.Properties.Item('Caption').Value = $definition.Caption
        $component.Properties.Item('Width').Value = [single]$definition.Width
        $component.Properties.Item('ShowModal').Value = $false
        $component.Properties.Item('StartUpPosition').Value = 0
        $designer = $component.Designer
        $designer.BackColor = 16447479
        $label = $designer.Controls.Add('Forms.Label.1','lblDocument',$true)
        $label.Left = 10; $label.Top = 8; $label.Width = $definition.Width - 30; $label.Height = 28
        $label.Caption = '当前文档'
        if ($toolbarKind -eq 'Final') {
            $snapshotLabel = $designer.Controls.Add('Forms.Label.1','lblSnapshot',$true)
            $snapshotLabel.Left = 10; $snapshotLabel.Top = 38; $snapshotLabel.Width = $definition.Width - 30; $snapshotLabel.Height = 30
            $snapshotLabel.Caption = '修订：—    批注：—    修订跟踪：—'
            $timeLabel = $designer.Controls.Add('Forms.Label.1','lblSnapshotTime',$true)
            $timeLabel.Left = 10; $timeLabel.Top = 74; $timeLabel.Width = $definition.Width - 94; $timeLabel.Height = 22
            $timeLabel.Caption = '状态读取时间：—'
            $refreshButton = $designer.Controls.Add('Forms.CommandButton.1','btnRefreshState',$true)
            $refreshButton.Left = $definition.Width - 80; $refreshButton.Top = 70; $refreshButton.Width = 54; $refreshButton.Height = 24
            $refreshButton.Caption = '刷新'; $refreshButton.Tag = 'refresh_final_state'; $refreshButton.TakeFocusOnClick = $false
        }
        $code = [Collections.Generic.List[string]]::new()
        $code.Add('Option Explicit')
        $code.Add('Private mUpdating As Boolean')
        $code.Add('Private mMenuRange As Range')
        if ($toolbarKind -eq 'Final') {
            $code.Add('Private mSnapshotDocument As Document')
            $code.Add('Private mSnapshotState As String, mSnapshotFailure As String')
            $code.Add('Private mSnapshotStale As Boolean, mSnapshotSaved As Boolean, mTracking As Boolean')
            $code.Add('Private mRevisionCount As Long, mCommentCount As Long, mReadAt As Date')
        }
        $code.Add('Private Sub UserForm_Initialize()')
        $code.Add('    Dim item As Object')
        $code.Add('    For Each item In Me.Controls: item.Font.Name = "Microsoft YaHei": item.Font.Size = 9: Next item')
        $code.Add('    On Error Resume Next')
        $code.Add('    Me.Left = CSng(GetSetting("RATools", "' + $toolbarKind + 'Toolbar", "Left", CStr(Application.Left + 30)))')
        $code.Add('    Me.Top = CSng(GetSetting("RATools", "' + $toolbarKind + 'Toolbar", "Top", CStr(Application.Top + 80)))')
        $code.Add('    On Error GoTo 0')
        $code.Add('    RefreshDocumentState False')
        $code.Add('End Sub')
        $top = if ($toolbarKind -eq 'Final') { 110 } else { 42 }
        $buttonIndex = 0
        $menuInitializers = [Collections.Generic.List[string]]::new()
        foreach ($row in $definition.Rows) {
            if ($row[0]) {
                $group = $designer.Controls.Add('Forms.Label.1', 'lblGroup' + $buttonIndex, $true)
                $group.Left = 10; $group.Top = $top; $group.Width = $definition.Width - 30; $group.Height = 18
                $group.Caption = $row[0]
                $top += 20
            }
            for ($column = 1; $column -lt $row.Count; $column++) {
                $parts = $row[$column] -split '\|',2
                $isMenu = $parts[1].StartsWith('menu:')
                $name = if ($isMenu) { 'cbo' + $parts[1].Substring(5) } else { 'btn' + $buttonIndex }
                $progId = if ($isMenu) { 'Forms.ComboBox.1' } else { 'Forms.CommandButton.1' }
                $control = $designer.Controls.Add($progId,$name,$true)
                $control.Left = 10 + ($column - 1) * (($definition.Width - 28) / $definition.Columns)
                $control.Top = $top; $control.Width = ($definition.Width - 28) / $definition.Columns - 5; $control.Height = 26
                $control.ControlTipText = if ($parts[1].StartsWith('style:')) { '应用 ' + $parts[1].Substring(6) + '；作用于当前选区或插入点。' } else { $parts[0] + '；作用范围和模式选择沿用原工具。' }
                if ($definition.ContainsKey('Tips') -and $definition.Tips.ContainsKey($parts[1])) { $control.ControlTipText = $definition.Tips[$parts[1]] }
                if ($isMenu) {
                    $control.Style = 2
                    $control.ListWidth = 280
                    $control.Tag = $parts[0]
                    $menuInitializers.Add('    ' + $name + '.AddItem ' + (VbaQuote $parts[0]))
                    $menuCases = [Collections.Generic.List[string]]::new()
                    $menuEntryIndex = 0
                    foreach ($entry in $definition.Menus[$parts[1].Substring(5)]) {
                        $menuEntryIndex++
                        $menuParts = $entry -split '\|',2
                        $menuInitializers.Add('    ' + $name + '.AddItem ' + (VbaQuote $menuParts[0]))
                        if ($menuParts[1]) { $menuCases.Add('        Case ' + $menuEntryIndex + ': actionKey = ' + (VbaQuote $menuParts[1])) }
                    }
                    $menuInitializers.Add('    ' + $name + '.ListIndex = 0')
                    $code.Add('Private Sub ' + $name + '_Change()')
                    $code.Add('    Dim actionKey As String')
                    $code.Add('    If mUpdating Or ' + $name + '.ListIndex < 1 Then Exit Sub')
                    $code.Add('    Select Case ' + $name + '.ListIndex')
                    foreach ($menuCase in $menuCases) { $code.Add($menuCase) }
                    $code.Add('    End Select')
                    $code.Add('    mUpdating = True: ' + $name + '.ListIndex = 0: mUpdating = False')
                    $code.Add('    On Error Resume Next')
                    $code.Add('    If Not mMenuRange Is Nothing Then')
                    $code.Add('        If Documents.Count > 0 Then')
                    $code.Add('            If mMenuRange.Document Is ActiveDocument Then mMenuRange.Select')
                    $code.Add('        End If')
                    $code.Add('    End If')
                    $code.Add('    Set mMenuRange = Nothing')
                    $code.Add('    On Error GoTo 0')
                    $code.Add('    If Len(actionKey) > 0 Then RunWorkflowAction actionKey')
                    $code.Add('End Sub')
                    $code.Add('Private Sub ' + $name + '_Enter()')
                    $code.Add('    On Error Resume Next')
                    $code.Add('    Set mMenuRange = Nothing')
                    $code.Add('    If Documents.Count > 0 Then Set mMenuRange = Selection.Range.Duplicate')
                    $code.Add('End Sub')
                    $code.Add('Private Sub ' + $name + '_DropButtonClick()')
                    $code.Add('    ' + $name + '_Enter')
                    $code.Add('End Sub')
                } else {
                    $control.Caption = $parts[0]
                    $control.Tag = $parts[1]
                    $control.TakeFocusOnClick = $false
                    $code.Add('Private Sub ' + $name + '_Click()')
                    $code.Add('    RunWorkflowAction ' + (VbaQuote $parts[1]))
                    $code.Add('End Sub')
                }
                $buttonIndex++
            }
            $top += 32
        }
        $component.Properties.Item('Height').Value = [single]($top + 30)
        $code.Add('Private Sub InitializeMenus()')
        $code.Add('    mUpdating = True')
        foreach ($line in $menuInitializers) { $code.Add($line) }
        $code.Add('    mUpdating = False')
        $code.Add('End Sub')
        if ($toolbarKind -eq 'Writing') {
        $code.Add('Public Sub RefreshDocumentState(ByVal busy As Boolean)')
        $code.Add('    Dim item As Object, available As Boolean, value As Long')
        $code.Add('    On Error GoTo Unavailable')
        $code.Add('    If cbostyles.ListCount = 0 Then InitializeMenus')
        $code.Add('    If Documents.Count = 0 Then lblDocument.Caption = "当前没有打开的文档": GoTo UpdateControls')
        $code.Add('    lblDocument.Caption = "当前文档：" & ActiveDocument.Name')
        $code.Add('    available = True')
        $code.Add('    value = Selection.ParagraphFormat.PageBreakBefore')
        $code.Add('    btn19.Caption = "段前分页：" & IIf(value = wdUndefined, "混合", IIf(value = 0, "关", "开"))')
        $code.Add('    GoTo UpdateControls')
        $code.Add('Unavailable:')
        $code.Add('    available = False: lblDocument.Caption = "当前文档不可用"')
        $code.Add('UpdateControls:')
        $code.Add('    For Each item In Me.Controls')
        $code.Add('        If TypeName(item) = "CommandButton" Or TypeName(item) = "ComboBox" Then item.Enabled = available And Not busy')
        $code.Add('    Next item')
        $code.Add('End Sub')
        } else {
            $finalStateCode = @'
Private Sub ResetSnapshot()
    Set mSnapshotDocument = Nothing
    mSnapshotState = "unread"
    mSnapshotStale = False
    mSnapshotFailure = ""
End Sub
Private Sub RenderSnapshot()
    Select Case mSnapshotState
        Case "read"
            lblSnapshot.Caption = IIf(mSnapshotStale, "上次读数（待刷新）：", "") & "修订：" & mRevisionCount & "    批注：" & mCommentCount & "    修订跟踪：" & IIf(mTracking, "开", "关")
            lblSnapshotTime.Caption = "读取时间：" & Format$(mReadAt, "hh:nn:ss") & IIf(mSnapshotStale, " · 待刷新", "")
        Case "failed"
            lblSnapshot.Caption = "状态读取失败：" & mSnapshotFailure
            lblSnapshotTime.Caption = "状态读取时间：—"
        Case Else
            lblSnapshot.Caption = "修订：—    批注：—    修订跟踪：—"
            lblSnapshotTime.Caption = "状态读取时间：—"
    End Select
End Sub
Public Sub ReadDocumentSnapshot()
    Dim values As Variant
    On Error GoTo Failed
    If Documents.Count = 0 Then Call ResetSnapshot: RenderSnapshot: Exit Sub
    Set mSnapshotDocument = ActiveDocument
    values = ReadWorkflowReviewSnapshot(mSnapshotDocument)
    mRevisionCount = CLng(values(0)): mCommentCount = CLng(values(1)): mTracking = CBool(values(2))
    mSnapshotSaved = mSnapshotDocument.Saved
    mReadAt = Now
    mSnapshotState = "read": mSnapshotStale = False: mSnapshotFailure = ""
    RenderSnapshot
    Exit Sub
Failed:
    mSnapshotState = "failed": mSnapshotStale = False
    mSnapshotFailure = Err.Description
    RenderSnapshot
End Sub
Public Sub RefreshDocumentState(ByVal busy As Boolean, Optional ByVal mayHaveEdited As Boolean = False)
    Dim item As Object, available As Boolean, needsRead As Boolean
    On Error GoTo Unavailable
    If Documents.Count = 0 Then
        lblDocument.Caption = "当前没有打开的文档"
        Call ResetSnapshot: RenderSnapshot
        GoTo UpdateControls
    End If
    lblDocument.Caption = "当前文档：" & ActiveDocument.Name & IIf(ActiveDocument.Saved, " · 已保存", " · 有未保存修改")
    available = True
    If mSnapshotDocument Is Nothing Then
        needsRead = True
    ElseIf Not (mSnapshotDocument Is ActiveDocument) Then
        ResetSnapshot
        needsRead = True
    End If
    If needsRead And Not busy Then ReadDocumentSnapshot
    If mSnapshotState = "read" Then
        If mayHaveEdited Or (mSnapshotSaved <> ActiveDocument.Saved) Then mSnapshotStale = True
    End If
    RenderSnapshot
    GoTo UpdateControls
Unavailable:
    available = False
    lblDocument.Caption = "当前文档不可用"
    Call ResetSnapshot: RenderSnapshot
UpdateControls:
    For Each item In Me.Controls
        If TypeName(item) = "CommandButton" Or TypeName(item) = "ComboBox" Then item.Enabled = available And Not busy
    Next item
End Sub
Private Sub btnRefreshState_Click()
    RunWorkflowAction "refresh_final_state"
End Sub
'@
            foreach ($stateLine in ($finalStateCode -split '\r?\n')) { $code.Add($stateLine) }
        }
        $code.Add('Private Sub UserForm_Activate(): RefreshWorkflowToolbarsIfVisible: End Sub')
        $code.Add('Public Sub SaveToolbarPosition()')
        $code.Add('    On Error Resume Next')
        $code.Add('    SaveSetting "RATools", "' + $toolbarKind + 'Toolbar", "Left", CStr(Me.Left)')
        $code.Add('    SaveSetting "RATools", "' + $toolbarKind + 'Toolbar", "Top", CStr(Me.Top)')
        $code.Add('End Sub')
        $code.Add('Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer): SaveToolbarPosition: End Sub')
        foreach ($item in $designer.Controls) {
            $code.Add('Private Sub ' + $item.Name + '_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)')
            $code.Add('    EnsureTooltipOnTop Me.Caption')
            $code.Add('End Sub')
        }
        $source = $code -join "`r`n"
        $source = $source.Replace('    RefreshDocumentState False', "    InitializeMenus`r`n    RefreshDocumentState False")
        $component.CodeModule.AddFromString($source)
        $component.Export((Join-Path $exportRoot "$formName.frm"))
        foreach ($extension in @('frm','frx')) {
            Copy-Item -LiteralPath (Join-Path $exportRoot "$formName.$extension") -Destination (Join-Path $repoRoot "userforms/$formName.$extension") -Force
        }
        Write-Output "PASS generated $formName"
    }
    finally {
        try {
            if ($null -ne $document) { $document.Close(0) }
        }
        finally {
            try { Stop-RAToolsWordSession -Word $word }
            finally { Move-RAToolsToRecycleBin -Path $exportRoot -SafeRoot ([IO.Path]::GetTempPath()) }
        }
    }
}
