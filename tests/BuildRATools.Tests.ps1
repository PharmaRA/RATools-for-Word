$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$modulePath = Join-Path $repoRoot "scripts\RATools.Build.psm1"

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Equal {
    param(
        $Expected,
        $Actual,
        [string]$Message
    )

    if ($Expected -ne $Actual) {
        throw "$Message Expected <$Expected>, got <$Actual>."
    }
}

function Assert-Throws {
    param(
        [scriptblock]$ScriptBlock,
        [string]$Message
    )

    try {
        & $ScriptBlock
    }
    catch {
        return
    }

    throw $Message
}

function New-TestRepo {
    $root = Join-Path ([System.IO.Path]::GetTempPath()) ("RATools_BuildTest_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $root | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $root "modules") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $root "class_modules") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $root "userforms") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $root "dotm\docProps") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $root "dotm\word") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $root "dist") | Out-Null

    Set-Content -LiteralPath (Join-Path $root "dotm\[Content_Types].xml") -Value "<Types />"
    Set-Content -LiteralPath (Join-Path $root "dotm\docProps\core.xml") -Value @"
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title></dc:title><dc:subject></dc:subject><dc:creator>Private User</dc:creator><cp:keywords></cp:keywords><dc:description></dc:description><cp:lastModifiedBy>Private User</cp:lastModifiedBy><cp:revision>7</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">2026-07-07T12:34:56Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-07T12:34:56Z</dcterms:modified></cp:coreProperties>
"@
    Set-Content -LiteralPath (Join-Path $root "dotm\docProps\app.xml") -Value @"
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>RATools.dotm</Template><Application>Microsoft Word</Application><Company>Private Company</Company><AppVersion>16.0000</AppVersion></Properties>
"@
    Set-Content -LiteralPath (Join-Path $root "dotm\word\document.xml") -Value "<document />"
    Set-Content -LiteralPath (Join-Path $root "dotm\word\vbaProject.bin") -Value "bin"

    return $root
}

function Read-ZipEntryText {
    param(
        [string]$PackagePath,
        [string]$EntryName
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
        $entry = $archive.GetEntry($EntryName)
        Assert-True ($null -ne $entry) "Missing zip entry: $EntryName"
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Run-Test {
    param(
        [string]$Name,
        [scriptblock]$Body
    )

    Write-Host "Running $Name"
    & $Body
}

if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "Expected build module not found: $modulePath"
}

Import-Module $modulePath -Force
Add-Type -AssemblyName System.IO.Compression.FileSystem

Run-Test "Get-RAToolsVbaSourceFiles returns importable files in stable order" {
    $root = New-TestRepo

    try {
        Set-Content -LiteralPath (Join-Path $root "modules\B_Module.bas") -Value "Attribute VB_Name = ""B_Module"""
        Set-Content -LiteralPath (Join-Path $root "modules\A_Module.bas") -Value "Attribute VB_Name = ""A_Module"""
        Set-Content -LiteralPath (Join-Path $root "class_modules\AppEvents.cls") -Value "VERSION 1.0 CLASS"
        Set-Content -LiteralPath (Join-Path $root "userforms\frmAbout.frm") -Value "VERSION 5.00"
        Set-Content -LiteralPath (Join-Path $root "userforms\frmAbout.frx") -Value "binary"

        $files = @(Get-RAToolsVbaSourceFiles -RepoRoot $root)
        $relativePaths = @($files | ForEach-Object { $_.RelativePath })

        Assert-Equal 4 $files.Count "Unexpected VBA source file count."
        Assert-Equal "modules\A_Module.bas" $relativePaths[0] "Standard modules should be sorted by name."
        Assert-Equal "modules\B_Module.bas" $relativePaths[1] "Standard modules should come before class modules."
        Assert-Equal "class_modules\AppEvents.cls" $relativePaths[2] "Class modules should come before user forms."
        Assert-Equal "userforms\frmAbout.frm" $relativePaths[3] "Only .frm should be imported for user forms."
    }
    finally {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Run-Test "New-RAToolsDotmFromDirectory writes package contents at zip root" {
    $root = New-TestRepo
    $outputPath = Join-Path $root "dist\RATools_test.dotm"

    try {
        New-RAToolsDotmFromDirectory -SourceDirectory (Join-Path $root "dotm") -DestinationPath $outputPath | Out-Null

        Assert-True (Test-Path -LiteralPath $outputPath) "Expected dotm package to be created."

        $archive = [System.IO.Compression.ZipFile]::OpenRead($outputPath)
        try {
            $entries = @($archive.Entries | ForEach-Object { $_.FullName })
            Assert-True ($entries -contains "[Content_Types].xml") "Package should contain [Content_Types].xml at root."
            Assert-True ($entries -contains "word/document.xml") "Package should contain word/document.xml at root."
            Assert-True ($entries -notcontains "dotm/[Content_Types].xml") "Package should not wrap contents in a dotm directory."
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Run-Test "Expand-RAToolsDotmToDirectory replaces stale target contents" {
    $root = New-TestRepo
    $outputPath = Join-Path $root "dist\RATools_test.dotm"
    $targetPath = Join-Path $root "expanded-dotm"

    try {
        New-RAToolsDotmFromDirectory -SourceDirectory (Join-Path $root "dotm") -DestinationPath $outputPath | Out-Null
        New-Item -ItemType Directory -Path $targetPath | Out-Null
        Set-Content -LiteralPath (Join-Path $targetPath "stale.txt") -Value "remove me"

        Expand-RAToolsDotmToDirectory -PackagePath $outputPath -TargetDirectory $targetPath -SafeRoot $root | Out-Null

        Assert-True (Test-Path -LiteralPath (Join-Path $targetPath "[Content_Types].xml")) "Expanded package should contain content types."
        Assert-True (-not (Test-Path -LiteralPath (Join-Path $targetPath "stale.txt"))) "Expansion should remove stale target files."
    }
    finally {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Run-Test "Assert-RAToolsPathInsideRoot rejects paths outside safe root" {
    $root = New-TestRepo
    $outside = Join-Path ([System.IO.Path]::GetTempPath()) ("RATools_Outside_" + [guid]::NewGuid().ToString("N"))

    try {
        Assert-Throws {
            Assert-RAToolsPathInsideRoot -Root $root -Path $outside -Description "outside test path"
        } "Path outside the safe root should be rejected."
    }
    finally {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Run-Test "Get-RAToolsLatestChangelogVersion reads the first changelog heading" {
    $root = New-TestRepo

    try {
        Set-Content -LiteralPath (Join-Path $root "CHANGELOG.md") -Value @"
# v9.8.7

- New release.

# v1.0.0

- Older release.
"@

        $version = Get-RAToolsLatestChangelogVersion -RepoRoot $root

        Assert-Equal "v9.8.7" $version "Latest changelog version should come from the first version heading."
    }
    finally {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Run-Test "Set-RAToolsAppVersion updates Mod_UpdateChecker APP_VERSION" {
    $root = New-TestRepo
    $modulePath = Join-Path $root "modules\Mod_UpdateChecker.bas"

    try {
        Set-Content -LiteralPath $modulePath -Value @"
Attribute VB_Name = "Mod_UpdateChecker"
Option Explicit

Private Const APP_VERSION As String = "v0.1.0"
Private Const GITHUB_REPOSITORY_URL As String = "https://github.com/PharmaRA/RATools-for-Word"
"@

        Set-RAToolsAppVersion -RepoRoot $root -Version "v9.8.7" | Out-Null

        $updatedText = Get-Content -LiteralPath $modulePath -Raw
        Assert-True ($updatedText.Contains('Private Const APP_VERSION As String = "v9.8.7"')) "APP_VERSION should be updated."
        Assert-True (-not $updatedText.Contains('Private Const APP_VERSION As String = "v0.1.0"')) "Old APP_VERSION should be removed."
        Assert-True ($updatedText.Contains('GITHUB_REPOSITORY_URL')) "Other module constants should be preserved."
    }
    finally {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Run-Test "Clear-RAToolsPackageMetadata removes personal document properties" {
    $root = New-TestRepo
    $outputPath = Join-Path $root "dist\RATools_private.dotm"

    try {
        New-RAToolsDotmFromDirectory -SourceDirectory (Join-Path $root "dotm") -DestinationPath $outputPath | Out-Null

        Clear-RAToolsPackageMetadata -PackagePath $outputPath | Out-Null

        [xml]$coreXml = Read-ZipEntryText -PackagePath $outputPath -EntryName "docProps/core.xml"
        [xml]$appXml = Read-ZipEntryText -PackagePath $outputPath -EntryName "docProps/app.xml"

        $ns = New-Object System.Xml.XmlNamespaceManager($coreXml.NameTable)
        $ns.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
        $ns.AddNamespace("dc", "http://purl.org/dc/elements/1.1/")
        $ns.AddNamespace("dcterms", "http://purl.org/dc/terms/")

        Assert-Equal "" $coreXml.SelectSingleNode("/cp:coreProperties/dc:creator", $ns).InnerText "Creator should be blank."
        Assert-Equal "" $coreXml.SelectSingleNode("/cp:coreProperties/cp:lastModifiedBy", $ns).InnerText "Last modified by should be blank."
        Assert-Equal "1" $coreXml.SelectSingleNode("/cp:coreProperties/cp:revision", $ns).InnerText "Revision should be reset."
        Assert-Equal "2000-01-01T00:00:00Z" $coreXml.SelectSingleNode("/cp:coreProperties/dcterms:created", $ns).InnerText "Created timestamp should be stable."
        Assert-Equal "2000-01-01T00:00:00Z" $coreXml.SelectSingleNode("/cp:coreProperties/dcterms:modified", $ns).InnerText "Modified timestamp should be stable."
        Assert-True (-not $coreXml.OuterXml.Contains("Private User")) "Core properties should not contain private user text."
        Assert-True (-not $appXml.OuterXml.Contains("Private Company")) "App properties should not contain private company text."
    }
    finally {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Write-Host "PASS BuildRATools.Tests"
