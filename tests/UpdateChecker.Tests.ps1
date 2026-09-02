$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "..\scripts\RATools.Build.psm1") -Force

# Engine_UpdateChecker 纯逻辑函数测试：版本比较与 JSON 字段提取。
# 需要 Word COM 执行 VBA，但不联网、不弹窗。

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $repoRoot "modules\Engine_UpdateChecker.bas"

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
    $probe.Name = "UpdateCheckerProbe"
    $probe.CodeModule.AddFromString(@'
Public Function ProbeCompare(ByVal current As String, ByVal latest As String) As Long
    ProbeCompare = CompareVersions(current, latest)
End Function

Public Function ProbeExtract(ByVal jsonText As String, ByVal keyName As String) As String
    ProbeExtract = ExtractJsonStringValue(jsonText, keyName)
End Function

Public Function ProbeNormalize(ByVal versionText As String) As String
    ProbeNormalize = NormalizeVersion(versionText)
End Function
'@)

    # --- CompareVersions：返回 1=有新版本, -1=本地更新, 0=相同 ---
    $compareCases = @(
        @{ Current = "v0.7.1"; Latest = "v0.7.2"; Expected = 1 },
        @{ Current = "v0.7.1"; Latest = "v0.8.0"; Expected = 1 },
        @{ Current = "v0.7.1"; Latest = "v1.0.0"; Expected = 1 },
        @{ Current = "v0.7.1"; Latest = "v0.7.1"; Expected = 0 },
        @{ Current = "v0.7.2"; Latest = "v0.7.1"; Expected = -1 },
        @{ Current = "v1.0.0"; Latest = "v0.9.9"; Expected = -1 },
        @{ Current = "0.7.1";  Latest = "V0.7.2"; Expected = 1 },   # 大小写与无前缀
        @{ Current = "v0.7";   Latest = "v0.7.1"; Expected = 1 },   # 缺段按 0
        @{ Current = "v0.7.1"; Latest = "";       Expected = -1 },  # 空串按 0.0.0
        @{ Current = "v0.7.1"; Latest = "vX.Y.Z"; Expected = -1 }   # 非数字按 0
    )
    foreach ($case in $compareCases) {
        $current = [string]$case.Current
        $latest = [string]$case.Latest
        $result = [long]$word.Run("ProbeCompare", [ref]$current, [ref]$latest)
        Assert-True ($result -eq $case.Expected) `
            "CompareVersions('$($case.Current)','$($case.Latest)') expected $($case.Expected), got $result"
    }
    Write-Host "CompareVersions: $($compareCases.Count) cases OK"

    # --- ExtractJsonStringValue ---
    $json = '{"tag_name": "v0.8.0", "html_url": "https://github.com/PharmaRA/RATools-for-Word/releases/tag/v0.8.0", "body": "line\"quoted\" end"}'
    $extractCases = @(
        @{ Key = "tag_name"; Expected = "v0.8.0" },
        @{ Key = "html_url"; Expected = "https://github.com/PharmaRA/RATools-for-Word/releases/tag/v0.8.0" },
        @{ Key = "body";     Expected = 'line\"quoted\" end' },  # 转义引号不截断
        @{ Key = "missing";  Expected = "" }
    )
    foreach ($case in $extractCases) {
        $key = [string]$case.Key
        $result = [string]$word.Run("ProbeExtract", [ref]$json, [ref]$key)
        Assert-True ($result -eq $case.Expected) `
            "ExtractJsonStringValue('$($case.Key)') expected <$($case.Expected)>, got <$result>"
    }
    Write-Host "ExtractJsonStringValue: $($extractCases.Count) cases OK"

    # --- NormalizeVersion ---
    $normalizeCases = @(
        @{ In = "v0.7.1"; Expected = "0.7.1" },
        @{ In = "V0.7.1"; Expected = "0.7.1" },
        @{ In = " 0.7.1 "; Expected = "0.7.1" },
        @{ In = ""; Expected = "" }
    )
    foreach ($case in $normalizeCases) {
        $inValue = [string]$case.In
        $result = [string]$word.Run("ProbeNormalize", [ref]$inValue)
        Assert-True ($result -eq $case.Expected) `
            "NormalizeVersion('$($case.In)') expected <$($case.Expected)>, got <$result>"
    }
    Write-Host "NormalizeVersion: $($normalizeCases.Count) cases OK"
}
finally {
    if ($null -ne $doc) { $doc.Saved = $true; $doc.Close(0) }
    Stop-RAToolsWordSession -Word $word
}

Write-Host "PASS UpdateChecker.Tests"
