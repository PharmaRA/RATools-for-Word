$ErrorActionPreference = "Stop"

# 全项目编译探针：把全部 VBA 源码导入新文档，为每个标准模块注入无参探针并执行。
# VBA 执行模块内任一过程都会强制编译该模块，语法错误会以异常形式暴露；
# 类模块与窗体通过 If False Then New 引用强制编译。
# 这是所有重构阶段的最低安全网：改动任何 .bas/.cls/.frm 后运行即可确认全工程可编译。

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Import-Module (Join-Path $repoRoot "scripts\RATools.Build.psm1") -Force

$word = $null
$doc = $null
$failedModules = @()

try {
    $sources = Get-RAToolsVbaSourceFiles -RepoRoot $repoRoot

    $word = Start-RAToolsWordSession

    $doc = $word.Documents.Add()
    $project = $doc.VBProject

    foreach ($source in $sources) {
        [void]$project.VBComponents.Import($source.Path)
    }

    # 分类组件：1=标准模块 2=类模块 3=窗体
    $stdModules = @()
    $classAndForms = @()
    foreach ($component in $project.VBComponents) {
        if ($component.Type -eq 1 -and $component.Name -ne "RAToolsProbeHost") {
            $stdModules += $component
        }
        elseif ($component.Type -eq 2 -or $component.Type -eq 3) {
            $classAndForms += $component
        }
    }

    # 标准模块：注入探针函数（执行即强制编译整个模块）
    foreach ($component in $stdModules) {
        $probeName = "RAToolsCompileProbe_" + $component.Name
        $component.CodeModule.AddFromString(@"
Public Function $probeName() As Boolean
    $probeName = True
End Function
"@)
    }

    # 类/窗体：注入引用探针模块（If False Then New 强制编译而不执行）
    $refLines = ""
    $i = 0
    foreach ($component in $classAndForms) {
        $i++
        $refLines += "    Dim ref$i As $($component.Name)`r`n"
        $refLines += "    If False Then Set ref$i = New $($component.Name)`r`n"
    }
    $probeHost = $project.VBComponents.Add(1)
    $probeHost.Name = "RAToolsProbeHost"
    $probeHost.CodeModule.AddFromString(@"
Public Function RAToolsCompileProbe_ClassesAndForms() As Boolean
$refLines    RAToolsCompileProbe_ClassesAndForms = True
End Function
"@)

    foreach ($component in $stdModules) {
        $probeName = "RAToolsCompileProbe_" + $component.Name
        Write-Host "Compiling $($component.Name)"
        try {
            $result = $word.Run($probeName)
            if ($result -ne $true) {
                $failedModules += "$($component.Name): probe returned $result"
            }
        }
        catch {
            $failedModules += "$($component.Name): $($_.Exception.Message)"
        }
    }

    Write-Host "Compiling class modules and userforms"
    try {
        $result = $word.Run("RAToolsCompileProbe_ClassesAndForms")
        if ($result -ne $true) {
            $failedModules += "ClassesAndForms: probe returned $result"
        }
    }
    catch {
        $failedModules += "ClassesAndForms: $($_.Exception.Message)"
    }
}
finally {
    if ($null -ne $doc) {
        $doc.Saved = $true
        $doc.Close(0)
    }
    Stop-RAToolsWordSession -Word $word
}

if ($failedModules.Count -gt 0) {
    Write-Host ""
    foreach ($failure in $failedModules) {
        Write-Host "COMPILE FAIL $failure" -ForegroundColor Red
    }
    throw "SourceCompile failed for $($failedModules.Count) module(s)"
}

Write-Host "PASS SourceCompile.Tests"
