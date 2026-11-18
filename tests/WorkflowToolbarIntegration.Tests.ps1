param([switch]$Visible, [string]$ProgressPath = '')
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'WorkflowToolbars.Tests.ps1') -Kind Joint -CompileOnly:(-not $Visible) -ProgressPath $ProgressPath
