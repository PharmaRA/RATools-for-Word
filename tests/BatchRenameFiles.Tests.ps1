$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

# Mod_BatchRenameFiles 文件名清洗规则测试（纯逻辑，不触碰文件系统）。

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $repoRoot "modules\Mod_BatchRenameFiles.bas"

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

    $probe = $doc.VBProject.VBComponents.Add(1)
    $probe.Name = "RenameProbe"
    $probe.CodeModule.AddFromString(@'
Public Function ProbeClean(ByVal baseName As String) As String
    Dim regExSpace As Object
    Dim regExClean As Object

    CreateRenameRegexes regExSpace, regExClean
    ProbeClean = CleanFileBaseName(baseName, regExSpace, regExClean)
End Function
'@)

    # 规则：小写化；字母/数字间空格→中划线；其余空格删除；
    # 非法字符→中划线；空结果回退 renamed-file
    $cases = @(
        @{ In = "My File 01";        Expected = "my-file-01" },
        @{ In = "Report Final";      Expected = "report-final" },
        @{ In = "中文 file";         Expected = "中文file" },
        @{ In = "中文 文档";         Expected = "中文文档" },
        @{ In = "file(1)";           Expected = "file-1-" },
        @{ In = "a b c";             Expected = "a-b-c" },
        @{ In = "UPPER_CASE-keep";   Expected = "upper_case-keep" },
        @{ In = "@#%";               Expected = "---" },
        @{ In = "混合 Mix 02 版";    Expected = "混合mix-02版" },
        @{ In = "  spaced  ";        Expected = "spaced" }
    )

    foreach ($case in $cases) {
        $inValue = [string]$case.In
        $result = [string]$word.Run("ProbeClean", [ref]$inValue)
        Assert-True ($result -eq $case.Expected) `
            "CleanFileBaseName('$($case.In)') expected <$($case.Expected)>, got <$result>"
    }
    Write-Host "CleanFileBaseName: $($cases.Count) cases OK"
}
finally {
    if ($null -ne $doc) { $doc.Saved = $true; $doc.Close(0) }
    Stop-RAToolsWordSession -Word $word
}

Write-Host "PASS BatchRenameFiles.Tests"
