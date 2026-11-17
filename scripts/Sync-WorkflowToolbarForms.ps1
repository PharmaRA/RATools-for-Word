#requires -version 5.1
[CmdletBinding()]
param([ValidateSet('Writing')][string[]]$Kind = @('Writing'))
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
        $code = [Collections.Generic.List[string]]::new()
        $code.Add('Option Explicit')
        $code.Add('Private mUpdating As Boolean')
        $code.Add('Private mMenuRange As Range')
        $code.Add('Private Sub UserForm_Initialize()')
        $code.Add('    Dim item As Object')
        $code.Add('    For Each item In Me.Controls: item.Font.Name = "Microsoft YaHei": item.Font.Size = 9: Next item')
        $code.Add('    On Error Resume Next')
        $code.Add('    Me.Left = CSng(GetSetting("RATools", "' + $toolbarKind + 'Toolbar", "Left", CStr(Application.Left + 30)))')
        $code.Add('    Me.Top = CSng(GetSetting("RATools", "' + $toolbarKind + 'Toolbar", "Top", CStr(Application.Top + 80)))')
        $code.Add('    On Error GoTo 0')
        $code.Add('    RefreshDocumentState False')
        $code.Add('End Sub')
        $top = 42
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
