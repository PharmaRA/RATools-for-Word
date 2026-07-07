Set-StrictMode -Version Latest

function Get-RAToolsFullPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "Path cannot be empty."
    }

    return [System.IO.Path]::GetFullPath($Path)
}

function Assert-RAToolsPathInsideRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [string]$Path,

        [string]$Description = "path"
    )

    $rootFull = (Get-RAToolsFullPath -Path $Root).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $pathFull = Get-RAToolsFullPath -Path $Path
    $comparison = [System.StringComparison]::OrdinalIgnoreCase

    if ($pathFull.Equals($rootFull, $comparison)) {
        return $pathFull
    }

    $rootPrefix = $rootFull + [System.IO.Path]::DirectorySeparatorChar
    if (-not $pathFull.StartsWith($rootPrefix, $comparison)) {
        throw "Refusing to operate on $Description outside safe root. Root: $rootFull Path: $pathFull"
    }

    return $pathFull
}

function Get-RAToolsProjectLayout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot
    )

    if (-not (Test-Path -LiteralPath $RepoRoot -PathType Container)) {
        throw "Repository root does not exist: $RepoRoot"
    }

    $resolvedRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

    return [pscustomobject]@{
        RepoRoot              = $resolvedRoot
        DotmDirectory         = Join-Path $resolvedRoot "dotm"
        ModulesDirectory      = Join-Path $resolvedRoot "modules"
        ClassModulesDirectory = Join-Path $resolvedRoot "class_modules"
        UserFormsDirectory    = Join-Path $resolvedRoot "userforms"
        TemplateDirectory     = Join-Path $resolvedRoot "template"
        DistDirectory         = Join-Path $resolvedRoot "dist"
        ScriptsDirectory      = Join-Path $resolvedRoot "scripts"
    }
}

function ConvertTo-RAToolsRelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $rootFull = (Get-RAToolsFullPath -Path $Root).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $pathFull = Assert-RAToolsPathInsideRoot -Root $rootFull -Path $Path -Description "relative path target"

    if ($pathFull.Equals($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ""
    }

    return $pathFull.Substring($rootFull.Length + 1)
}

function Get-RAToolsVbaSourceFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot
    )

    $layout = Get-RAToolsProjectLayout -RepoRoot $RepoRoot
    $groups = @(
        [pscustomobject]@{
            Directory = $layout.ModulesDirectory
            Filter    = "*.bas"
            Type      = "StandardModule"
            Order     = 0
        },
        [pscustomobject]@{
            Directory = $layout.ClassModulesDirectory
            Filter    = "*.cls"
            Type      = "ClassModule"
            Order     = 1
        },
        [pscustomobject]@{
            Directory = $layout.UserFormsDirectory
            Filter    = "*.frm"
            Type      = "UserForm"
            Order     = 2
        }
    )

    $sourceFiles = foreach ($group in $groups) {
        if (-not (Test-Path -LiteralPath $group.Directory -PathType Container)) {
            continue
        }

        Get-ChildItem -LiteralPath $group.Directory -Filter $group.Filter -File |
            Sort-Object -Property Name |
            ForEach-Object {
                [pscustomobject]@{
                    Path         = $_.FullName
                    RelativePath = ConvertTo-RAToolsRelativePath -Root $layout.RepoRoot -Path $_.FullName
                    Type         = $group.Type
                    Order        = $group.Order
                    Name         = $_.Name
                }
            }
    }

    return @($sourceFiles | Sort-Object -Property Order, Name)
}

function Test-RAToolsDotmDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$ThrowOnFailure
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        if ($ThrowOnFailure) {
            throw "Missing dotm directory: $Path"
        }
        return $false
    }

    $requiredPaths = @(
        "[Content_Types].xml",
        "word\document.xml",
        "word\vbaProject.bin"
    )

    foreach ($relativePath in $requiredPaths) {
        $candidate = Join-Path $Path $relativePath
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            if ($ThrowOnFailure) {
                throw "Dotm directory is missing required package part: $relativePath"
            }
            return $false
        }
    }

    return $true
}

function New-RAToolsDotmFromDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceDirectory,

        [Parameter(Mandatory)]
        [string]$DestinationPath
    )

    $sourceFull = (Resolve-Path -LiteralPath $SourceDirectory).Path
    [void](Test-RAToolsDotmDirectory -Path $sourceFull -ThrowOnFailure)

    $destinationFull = Get-RAToolsFullPath -Path $DestinationPath
    $destinationParent = Split-Path -Parent $destinationFull

    if (-not (Test-Path -LiteralPath $destinationParent -PathType Container)) {
        New-Item -ItemType Directory -Path $destinationParent | Out-Null
    }

    if (Test-Path -LiteralPath $destinationFull) {
        Remove-Item -LiteralPath $destinationFull -Force
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::Open(
        $destinationFull,
        [System.IO.Compression.ZipArchiveMode]::Create
    )

    try {
        Get-ChildItem -LiteralPath $sourceFull -Recurse -File |
            Sort-Object -Property FullName |
            ForEach-Object {
                $entryName = (ConvertTo-RAToolsRelativePath -Root $sourceFull -Path $_.FullName) -replace "\\", "/"
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                    $archive,
                    $_.FullName,
                    $entryName,
                    [System.IO.Compression.CompressionLevel]::Optimal
                ) | Out-Null
            }
    }
    finally {
        $archive.Dispose()
    }

    return $destinationFull
}

function Expand-RAToolsDotmToDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PackagePath,

        [Parameter(Mandatory)]
        [string]$TargetDirectory,

        [string]$SafeRoot
    )

    if (-not (Test-Path -LiteralPath $PackagePath -PathType Leaf)) {
        throw "Missing dotm package: $PackagePath"
    }

    $packageFull = (Resolve-Path -LiteralPath $PackagePath).Path
    $targetFull = Get-RAToolsFullPath -Path $TargetDirectory

    if (-not [string]::IsNullOrWhiteSpace($SafeRoot)) {
        $targetFull = Assert-RAToolsPathInsideRoot -Root $SafeRoot -Path $targetFull -Description "dotm extraction target"
    }

    if (Test-Path -LiteralPath $targetFull) {
        Remove-Item -LiteralPath $targetFull -Recurse -Force
    }

    New-Item -ItemType Directory -Path $targetFull | Out-Null

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($packageFull, $targetFull)
    [void](Test-RAToolsDotmDirectory -Path $targetFull -ThrowOnFailure)

    return $targetFull
}

function Get-RAToolsLatestChangelogVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot
    )

    $layout = Get-RAToolsProjectLayout -RepoRoot $RepoRoot
    $changelogPath = Join-Path $layout.RepoRoot "CHANGELOG.md"

    if (-not (Test-Path -LiteralPath $changelogPath -PathType Leaf)) {
        throw "Missing changelog file: $changelogPath"
    }

    $changelogText = Get-Content -LiteralPath $changelogPath -Raw
    $match = [Regex]::Match($changelogText, "(?m)^\s*#\s+(v\d+(?:\.\d+){1,3}(?:[-+][0-9A-Za-z.-]+)?)\s*$")

    if (-not $match.Success) {
        throw "Could not find a version heading like '# v0.6.4' in $changelogPath"
    }

    return $match.Groups[1].Value
}

function Set-RAToolsAppVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot,

        [Parameter(Mandatory)]
        [string]$Version
    )

    if ([string]::IsNullOrWhiteSpace($Version)) {
        throw "App version cannot be empty."
    }

    $layout = Get-RAToolsProjectLayout -RepoRoot $RepoRoot
    $modulePath = Join-Path $layout.ModulesDirectory "Mod_UpdateChecker.bas"

    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
        throw "Missing update checker module: $modulePath"
    }

    $moduleText = Get-Content -LiteralPath $modulePath -Raw
    $versionPattern = "(?m)^(?<prefix>\s*Private\s+Const\s+APP_VERSION\s+As\s+String\s*=\s*)""(?<version>[^""]*)"""
    $versionRegex = [Regex]::new($versionPattern)
    $match = $versionRegex.Match($moduleText)

    if (-not $match.Success) {
        throw "Could not find APP_VERSION constant in $modulePath"
    }

    $updatedText = $versionRegex.Replace(
        $moduleText,
        {
            param($currentMatch)
            $currentMatch.Groups["prefix"].Value + """" + $Version + """"
        },
        1
    )

    if ($updatedText -ne $moduleText) {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($modulePath, $updatedText, $utf8NoBom)
    }

    return $modulePath
}

function ConvertTo-RAToolsXmlText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [xml]$XmlDocument
    )

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.OmitXmlDeclaration = $true

    $builder = New-Object System.Text.StringBuilder
    $writer = [System.Xml.XmlWriter]::Create($builder, $settings)

    try {
        $XmlDocument.Save($writer)
    }
    finally {
        $writer.Dispose()
    }

    return $builder.ToString()
}

function Set-RAToolsXmlNodeText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [xml]$XmlDocument,

        [Parameter(Mandatory)]
        [string]$XPath,

        [Parameter(Mandatory)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager,

        [AllowEmptyString()]
        [string]$Value
    )

    $node = $XmlDocument.SelectSingleNode($XPath, $NamespaceManager)
    if ($null -ne $node) {
        $node.InnerText = $Value
    }
}

function Set-RAToolsXmlLocalNameText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [xml]$XmlDocument,

        [Parameter(Mandatory)]
        [string]$LocalName,

        [AllowEmptyString()]
        [string]$Value
    )

    $node = $XmlDocument.SelectSingleNode("/*[local-name()='Properties']/*[local-name()='$LocalName']")
    if ($null -ne $node) {
        $node.InnerText = $Value
    }
}

function ConvertTo-RAToolsSanitizedCorePropertiesXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$XmlText
    )

    [xml]$xml = $XmlText
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
    $ns.AddNamespace("dc", "http://purl.org/dc/elements/1.1/")
    $ns.AddNamespace("dcterms", "http://purl.org/dc/terms/")

    Set-RAToolsXmlNodeText -XmlDocument $xml -XPath "/cp:coreProperties/dc:creator" -NamespaceManager $ns -Value ""
    Set-RAToolsXmlNodeText -XmlDocument $xml -XPath "/cp:coreProperties/cp:lastModifiedBy" -NamespaceManager $ns -Value ""
    Set-RAToolsXmlNodeText -XmlDocument $xml -XPath "/cp:coreProperties/cp:revision" -NamespaceManager $ns -Value "1"
    Set-RAToolsXmlNodeText -XmlDocument $xml -XPath "/cp:coreProperties/dcterms:created" -NamespaceManager $ns -Value "2000-01-01T00:00:00Z"
    Set-RAToolsXmlNodeText -XmlDocument $xml -XPath "/cp:coreProperties/dcterms:modified" -NamespaceManager $ns -Value "2000-01-01T00:00:00Z"

    return ConvertTo-RAToolsXmlText -XmlDocument $xml
}

function ConvertTo-RAToolsSanitizedAppPropertiesXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$XmlText
    )

    [xml]$xml = $XmlText
    Set-RAToolsXmlLocalNameText -XmlDocument $xml -LocalName "Company" -Value ""
    Set-RAToolsXmlLocalNameText -XmlDocument $xml -LocalName "Manager" -Value ""

    return ConvertTo-RAToolsXmlText -XmlDocument $xml
}

function Update-RAToolsZipEntryText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory)]
        [string]$EntryName,

        [Parameter(Mandatory)]
        [scriptblock]$Transform
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        return
    }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    try {
        $currentText = $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }

    $updatedText = & $Transform $currentText
    $entry.Delete()

    $newEntry = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter(
        $newEntry.Open(),
        (New-Object System.Text.UTF8Encoding($false))
    )

    try {
        $writer.Write($updatedText)
    }
    finally {
        $writer.Dispose()
    }
}

function Clear-RAToolsPackageMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PackagePath
    )

    if (-not (Test-Path -LiteralPath $PackagePath -PathType Leaf)) {
        throw "Missing dotm package: $PackagePath"
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $packageFull = (Resolve-Path -LiteralPath $PackagePath).Path
    $archive = [System.IO.Compression.ZipFile]::Open(
        $packageFull,
        [System.IO.Compression.ZipArchiveMode]::Update
    )

    try {
        Update-RAToolsZipEntryText -Archive $archive -EntryName "docProps/core.xml" -Transform {
            param([string]$text)
            ConvertTo-RAToolsSanitizedCorePropertiesXml -XmlText $text
        }

        Update-RAToolsZipEntryText -Archive $archive -EntryName "docProps/app.xml" -Transform {
            param([string]$text)
            ConvertTo-RAToolsSanitizedAppPropertiesXml -XmlText $text
        }
    }
    finally {
        $archive.Dispose()
    }

    return $packageFull
}

Export-ModuleMember -Function @(
    "Assert-RAToolsPathInsideRoot",
    "Get-RAToolsProjectLayout",
    "Get-RAToolsVbaSourceFiles",
    "Get-RAToolsLatestChangelogVersion",
    "Set-RAToolsAppVersion",
    "Test-RAToolsDotmDirectory",
    "New-RAToolsDotmFromDirectory",
    "Expand-RAToolsDotmToDirectory",
    "Clear-RAToolsPackageMetadata"
)
