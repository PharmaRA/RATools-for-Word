#requires -version 5.1

[CmdletBinding()]
param(
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
    $OutputPath = Join-Path $repoRoot "dotm\customUI\images\QuickToolbar.png"
}

function New-RoundedRectanglePath {
    param(
        [single]$X,
        [single]$Y,
        [single]$Width,
        [single]$Height,
        [single]$Radius
    )

    $diameter = $Radius * 2
    $path = New-Object Drawing.Drawing2D.GraphicsPath
    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
    $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-RoundPen {
    param(
        [Drawing.Color]$Color,
        [single]$Width
    )

    $pen = New-Object Drawing.Pen $Color, $Width
    $pen.LineJoin = [Drawing.Drawing2D.LineJoin]::Round
    $pen.StartCap = [Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [Drawing.Drawing2D.LineCap]::Round
    return $pen
}

Add-Type -AssemblyName System.Drawing

$scale = 4
$working = New-Object Drawing.Bitmap (48 * $scale), (48 * $scale), ([Drawing.Imaging.PixelFormat]::Format32bppArgb)
$graphics = [Drawing.Graphics]::FromImage($working)
$outputBitmap = $null

$outlineColor = [Drawing.Color]::FromArgb(255, 34, 34, 34)
$backFillColor = [Drawing.Color]::FromArgb(255, 218, 218, 218)
$panelFillColor = [Drawing.Color]::FromArgb(255, 137, 137, 137)
$accentColor = [Drawing.Color]::FromArgb(255, 226, 24, 31)
$buttonColor = [Drawing.Color]::FromArgb(255, 39, 39, 39)

$outlinePen = New-RoundPen -Color $outlineColor -Width 3.4
$backLinePen = New-RoundPen -Color $outlineColor -Width 2.2
$accentPen = New-RoundPen -Color $accentColor -Width 3.4
$backBrush = New-Object Drawing.SolidBrush $backFillColor
$panelBrush = New-Object Drawing.SolidBrush $panelFillColor
$buttonBrush = New-Object Drawing.SolidBrush $buttonColor

try {
    $graphics.Clear([Drawing.Color]::Transparent)
    $graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.ScaleTransform($scale, $scale)

    $backWindow = New-RoundedRectanglePath -X 14 -Y 4 -Width 30 -Height 29 -Radius 3.5
    try {
        $graphics.FillPath($backBrush, $backWindow)
        $graphics.DrawPath($outlinePen, $backWindow)
        $graphics.DrawLine($backLinePen, 18, 11, 40, 11)
    }
    finally {
        $backWindow.Dispose()
    }

    $floatingPanel = New-RoundedRectanglePath -X 4 -Y 12 -Width 34 -Height 32 -Radius 4
    try {
        $graphics.FillPath($panelBrush, $floatingPanel)
        $graphics.DrawPath($outlinePen, $floatingPanel)
    }
    finally {
        $floatingPanel.Dispose()
    }

    $graphics.DrawLine($accentPen, 10, 19, 32, 19)

    $buttonPositions = @(
        @(8, 24.5), @(15, 24.5), @(22, 24.5), @(29, 24.5),
        @(8, 34), @(15, 34), @(22, 34), @(29, 34)
    )
    foreach ($position in $buttonPositions) {
        $buttonPath = New-RoundedRectanglePath -X $position[0] -Y $position[1] -Width 5 -Height 5.5 -Radius 1
        try {
            $graphics.FillPath($buttonBrush, $buttonPath)
        }
        finally {
            $buttonPath.Dispose()
        }
    }

    $outputDirectory = Split-Path -Parent ([IO.Path]::GetFullPath($OutputPath))
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }

    $outputBitmap = New-Object Drawing.Bitmap 48, 48, ([Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $outputGraphics = [Drawing.Graphics]::FromImage($outputBitmap)
    try {
        $outputGraphics.Clear([Drawing.Color]::Transparent)
        $outputGraphics.CompositingMode = [Drawing.Drawing2D.CompositingMode]::SourceCopy
        $outputGraphics.CompositingQuality = [Drawing.Drawing2D.CompositingQuality]::HighQuality
        $outputGraphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $outputGraphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $outputGraphics.DrawImage($working, 0, 0, 48, 48)
    }
    finally {
        $outputGraphics.Dispose()
    }

    $outputBitmap.Save([IO.Path]::GetFullPath($OutputPath), [Drawing.Imaging.ImageFormat]::Png)
    Write-Output "Generated quick toolbar icon: $([IO.Path]::GetFullPath($OutputPath))"
}
finally {
    if ($null -ne $outputBitmap) { $outputBitmap.Dispose() }
    $buttonBrush.Dispose()
    $panelBrush.Dispose()
    $backBrush.Dispose()
    $accentPen.Dispose()
    $backLinePen.Dispose()
    $outlinePen.Dispose()
    $graphics.Dispose()
    $working.Dispose()
}
