#requires -version 5.1

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$Version = "local",
    [string]$AppVersion,
    [string]$OutputPath,
    [string]$BaseDotmPath,
    [switch]$NoSyncDotmDirectory,
    [switch]$SkipVbaImport,
    [switch]$KeepWorkDirectory
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $PSCommandPath
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $scriptRoot "..")).Path
}

Import-Module (Join-Path $scriptRoot "RATools.Build.psm1") -Force

$wdFormatXMLTemplateMacroEnabled = 15
$vbextCtDocument = 100

function Release-ComObjectSafe {
    param($ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Import-RAToolsVbaSourcesIntoDotm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DotmPath,

        [Parameter(Mandatory)]
        [string]$SourceRepoRoot
    )

    $sourceFiles = @(Get-RAToolsVbaSourceFiles -RepoRoot $SourceRepoRoot)
    if ($sourceFiles.Count -eq 0) {
        throw "No VBA source files found under modules, class_modules, or userforms."
    }

    $word = $null
    $document = $null
    $components = $null

    try {
        Write-Host "Starting Word automation"
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 1

        Write-Host "Opening dotm package: $DotmPath"
        $document = $word.Documents.Open($DotmPath, $false, $false, $false)
        $components = $document.VBProject.VBComponents

        $removeList = @()
        foreach ($component in $components) {
            if ([int]$component.Type -ne $vbextCtDocument) {
                $removeList += $component
            }
        }

        foreach ($component in $removeList) {
            Write-Host "Removing VBA component: $($component.Name)"
            $components.Remove($component)
        }

        foreach ($sourceFile in $sourceFiles) {
            Write-Host "Importing $($sourceFile.RelativePath)"
            [void]$components.Import($sourceFile.Path)
        }

        Write-Host "Saving updated dotm"
        $document.SaveAs2($DotmPath, $wdFormatXMLTemplateMacroEnabled)
        $document.Saved = $true
    }
    catch {
        throw "Could not update VBA project in $DotmPath. Enable 'Trust access to the VBA project object model' in Word Trust Center, close any open RATools template, then retry. $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $document) {
            $document.Close($false) | Out-Null
        }
        if ($null -ne $word) {
            $word.Quit() | Out-Null
        }

        Release-ComObjectSafe $components
        Release-ComObjectSafe $document
        Release-ComObjectSafe $word
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-SafeVersionName {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "local"
    }

    return ($Value -replace "[^A-Za-z0-9._-]", "-")
}

$layout = Get-RAToolsProjectLayout -RepoRoot $RepoRoot
[void](Test-RAToolsDotmDirectory -Path $layout.DotmDirectory -ThrowOnFailure)

if (-not $SkipVbaImport -and [string]::IsNullOrWhiteSpace($AppVersion)) {
    $AppVersion = Get-RAToolsLatestChangelogVersion -RepoRoot $layout.RepoRoot
}

$safeVersion = Get-SafeVersionName -Value $Version
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $layout.DistDirectory "RATools_$safeVersion.dotm"
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$workRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("RATools_Build_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workRoot | Out-Null

try {
    $basePackagePath = Join-Path $workRoot "base.dotm"
    $workingPackagePath = Join-Path $workRoot "working.dotm"

    if ([string]::IsNullOrWhiteSpace($BaseDotmPath)) {
        Write-Host "Creating base package from dotm directory"
        New-RAToolsDotmFromDirectory -SourceDirectory $layout.DotmDirectory -DestinationPath $basePackagePath | Out-Null
    }
    else {
        if (-not (Test-Path -LiteralPath $BaseDotmPath -PathType Leaf)) {
            throw "Base dotm file does not exist: $BaseDotmPath"
        }

        Copy-Item -LiteralPath $BaseDotmPath -Destination $basePackagePath -Force
    }

    Copy-Item -LiteralPath $basePackagePath -Destination $workingPackagePath -Force

    if ($SkipVbaImport) {
        Write-Host "Skipping Word VBA import; packaging current dotm content only"
    }
    else {
        Write-Host "Syncing APP_VERSION to $AppVersion"
        Set-RAToolsAppVersion -RepoRoot $layout.RepoRoot -Version $AppVersion | Out-Null
        Import-RAToolsVbaSourcesIntoDotm -DotmPath $workingPackagePath -SourceRepoRoot $layout.RepoRoot
    }

    Write-Host "Clearing document metadata"
    Clear-RAToolsPackageMetadata -PackagePath $workingPackagePath | Out-Null

    if (-not $NoSyncDotmDirectory) {
        Write-Host "Syncing dotm directory from generated package"
        Expand-RAToolsDotmToDirectory -PackagePath $workingPackagePath -TargetDirectory $layout.DotmDirectory -SafeRoot $layout.RepoRoot | Out-Null
    }

    if (Test-Path -LiteralPath $outputFullPath) {
        Remove-Item -LiteralPath $outputFullPath -Force
    }

    Copy-Item -LiteralPath $workingPackagePath -Destination $outputFullPath -Force
    Write-Host "Built dotm: $outputFullPath"
}
finally {
    if ($KeepWorkDirectory) {
        Write-Host "Keeping work directory: $workRoot"
    }
    else {
        Remove-Item -LiteralPath $workRoot -Recurse -Force
    }
}
