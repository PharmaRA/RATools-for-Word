$ErrorActionPreference = "Stop"

# PowerShell 脚本编码契约（无 COM）：含非 ASCII 字符的 .ps1/.psm1/.psd1 必须是
# 带 BOM 的 UTF-8。
#
# Windows PowerShell 5.1 对无 BOM 文件按系统 ANSI 代码页解码。在中文系统（936）上，
# UTF-8 中文注释被按双字节 GBK 拆读，行尾那个落单的前导字节会与 CRLF 的 0x0A 配成
# 一个字符，把换行吞掉——紧跟其后的那行代码于是变成注释的一部分，静默失效。
# scripts/RATools.Build.psm1 就因此让 Set-RAToolsAppVersion 报
# "无法检索变量 $gbkEncoding"，而同样的文件在 CI（ANSI 为 1252，单字节解码不吞换行）
# 上一切正常，所以这类问题只在中文机器上暴露。BOM 让解码不再依赖系统代码页。
#
# 注意：本契约只管 PowerShell 脚本。VBA 源码（.bas/.cls/.frm）走的是相反的
# GBK(936) 约定，见 docs/vba-source-encoding.md。

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

Write-Output "Running PowerShell script encoding checks"

# 只检查仓库内的源码：.git 与 .claude（worktree 副本、被 .gitignore 忽略）里的
# 脚本不是本仓库要维护的内容，历史副本的编码不该让测试失败。
$scripts = Get-ChildItem -LiteralPath $repoRoot -Recurse -File |
    Where-Object { @(".ps1", ".psm1", ".psd1") -contains $_.Extension } |
    Where-Object { $_.FullName -notmatch "\\\.(git|claude)\\" } |
    Sort-Object FullName

Assert-True ($scripts.Count -ge 20) "Should find the repository's PowerShell files; found $($scripts.Count)."

$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
$checked = 0

foreach ($script in $scripts) {
    $relative = $script.FullName.Substring($repoRoot.Length + 1)
    $bytes = [System.IO.File]::ReadAllBytes($script.FullName)

    $hasNonAscii = $false
    foreach ($byte in $bytes) {
        if ($byte -gt 127) {
            $hasNonAscii = $true
            break
        }
    }

    # 纯 ASCII 脚本在任何代码页下解码结果一致，不要求 BOM。
    if (-not $hasNonAscii) { continue }

    $checked++

    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    Assert-True $hasBom "$relative contains non-ASCII text and must be saved as UTF-8 with BOM, or Windows PowerShell 5.1 on a Chinese (936) system mis-decodes it and can swallow line breaks."

    $valid = $true
    try {
        $null = $strictUtf8.GetString($bytes)
    }
    catch {
        $valid = $false
    }
    Assert-True $valid "$relative is not valid UTF-8; a PowerShell script must not be saved as GBK (that convention applies to VBA sources only)."
}

Assert-True ($checked -ge 10) "Should have checked the non-ASCII PowerShell files; checked $checked."

Write-Output "Checked $checked non-ASCII PowerShell file(s) of $($scripts.Count)"
Write-Output "PASS ScriptEncoding.Tests ($assertions assertions)"
