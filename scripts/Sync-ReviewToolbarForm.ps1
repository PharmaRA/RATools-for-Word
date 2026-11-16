#requires -version 5.1
[CmdletBinding()]
param([string]$RepoRoot)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($RepoRoot)) { $RepoRoot = Split-Path -Parent $PSScriptRoot }
Import-Module (Join-Path $PSScriptRoot 'RATools.Build.psm1') -Force
$formPath = Join-Path $RepoRoot 'userforms/frmReviewToolbar.frm'
$gbk = [Text.Encoding]::GetEncoding(936)
$formText = [IO.File]::ReadAllText($formPath, $gbk)
$codeStart = $formText.IndexOf('Option Explicit')
if ($codeStart -lt 0) { throw 'Missing form code.' }
$formCode = $formText.Substring($codeStart)
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('RATools_ReviewForm_' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($workRoot)
$backupRoot = Join-Path $RepoRoot ('backups/review-form-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($backupRoot)
foreach ($sourcePath in @($formPath, [IO.Path]::ChangeExtension($formPath, '.frx'))) {
    if (Test-Path -LiteralPath $sourcePath) { Copy-Item -LiteralPath $sourcePath -Destination $backupRoot }
}
$word = $null
$document = $null
$component = $null
$designer = $null
function Add-Label([string]$Name, [string]$Caption, [single]$Top, [single]$Height = 16) {
    $item = $designer.Controls.Add('Forms.Label.1', $Name, $true)
    $item.Caption = $Caption
    $item.Left = 12
    $item.Top = $Top
    $item.Width = 202
    $item.Height = $Height
    $item.BackStyle = 0
    return $item
}
function Add-Button([string]$Name, [string]$Caption, [single]$Top, [string]$Tip) {
    $item = $designer.Controls.Add('Forms.CommandButton.1', $Name, $true)
    $item.Caption = $Caption
    $item.ControlTipText = $Tip
    $item.Left = 12
    $item.Top = $Top
    $item.Width = 202
    $item.Height = 24
    $item.TakeFocusOnClick = $false
    $item.TabIndex = $script:buttonIndex
    $script:buttonIndex++
}
try {
    $word = Start-RAToolsWordSession
    $document = $word.Documents.Add()
    $component = $document.VBProject.VBComponents.Add(3)
    $component.Name = 'frmReviewToolbar'
    $component.Properties.Item('Caption').Value = 'RATools 审阅工具'
    $component.Properties.Item('Width').Value = 238
    $component.Properties.Item('Height').Value = 425
    $component.Properties.Item('ShowModal').Value = $false
    $component.Properties.Item('StartUpPosition').Value = 0
    $designer = $component.Designer
    $designer.BackColor = 16447479
    [void](Add-Label 'lblDocument' '当前文档' 10 28)
    [void](Add-Label 'lblDaily' '日常审阅 · 当前文档 / 选区' 44)
    $script:buttonIndex = 0
    Add-Button 'btnTrackChanges' '修订跟踪：不可用' 64 '切换当前文档修订跟踪；关闭跟踪不清除已有修订。'
    Add-Button 'btnNewComment' '新建批注' 92 '使用正文中原有选区，打开 Word 批注编辑。'
    Add-Button 'btnReviewerFilter' '按审阅者筛选' 120 '仅改变当前文档标记显示；打开已有审阅者筛选面板。'
    [void](Add-Label 'lblFormat' '格式处理 · 当前文档全文' 156)
    Add-Button 'btnAcceptFormat' '接受格式修订' 176 '处理当前文档全文；快速路径会保存、关闭并重新打开文件。'
    [void](Add-Label 'lblCheck' '定稿前检查 · 选择检测模式' 212)
    Add-Button 'btnDetectHighlights' '检测高亮内容' 232 '当前文档 / 多文件 / 文件夹；检测而不清除高亮。'
    Add-Button 'btnDetectClean' '检测 clean 版本' 260 '只读检查修订、批注和跟踪开关；高亮不纳入 clean 判定。'
    [void](Add-Label 'lblClean' '清稿处理 · 确认意见后使用' 300)
    Add-Button 'btnClean' '接受全部修订并删除批注' 322 '接受全部修订、删除批注、关闭跟踪；保留原模式选择，批量模式写回原文件。'
    [void](Add-Label 'lblWarning' '关闭跟踪；批量模式会写回原文件。' 352 30)
    # Ensure newly generated controls also keep their own tooltips above the form.
    foreach ($controlItem in $designer.Controls) {
        $handlerName = [string]$controlItem.Name + '_MouseMove'
        if (-not $formCode.Contains('Sub ' + $handlerName + '(')) {
            $formCode += "`r`nPrivate Sub ${handlerName}(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)`r`n    RefreshReviewToolbarTooltip`r`nEnd Sub`r`n"
        }
    }
    $component.CodeModule.AddFromString($formCode)
    $exportPath = Join-Path $workRoot 'frmReviewToolbar.frm'
    $component.Export($exportPath)
    $resourcePath = [IO.Path]::ChangeExtension($exportPath, '.frx')
    if (-not (Test-Path -LiteralPath $resourcePath)) { throw 'Missing exported FRX.' }
    Copy-Item -LiteralPath $exportPath -Destination $formPath -Force
    Copy-Item -LiteralPath $resourcePath -Destination ([IO.Path]::ChangeExtension($formPath, '.frx')) -Force
    Write-Output 'PASS generated review toolbar: seven buttons, four groups.'
}
finally {
    if ($null -ne $document) { $document.Close(0) }
    foreach ($comItem in @($designer, $component, $document)) {
        if ($null -ne $comItem -and [Runtime.InteropServices.Marshal]::IsComObject($comItem)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comItem)
        }
    }
    Stop-RAToolsWordSession -Word $word
    Move-RAToolsToRecycleBin -Path $workRoot -SafeRoot ([IO.Path]::GetTempPath())
}
