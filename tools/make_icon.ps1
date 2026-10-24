# Regenerate the 64x64 application icon from the AI-generated source image.
#
# Source : <repo>\Calendar.png                       (1536x1536 RGBA, transparent background)
# Output : <repo>\OutlookDateJump\CalendarIcon.png    (64x64, alpha kept, margins trimmed)
#
# The generated icon is used in two places:
#   * Ribbon button    - RibbonImageHelper.GetIcon("CalendarIcon")  (Ribbon1.xml: image="CalendarIcon")
#   * Dialog title bar - DatePickerWindow, shown at 32x32 inside the rounded plate
#
# Paths are resolved relative to this script, so it runs on any machine
# (no hardcoded drive / user folder).

Add-Type -AssemblyName System.Drawing

$repo = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $repo 'Calendar.png'
$dst  = Join-Path $repo 'OutlookDateJump\CalendarIcon.png'

if (-not (Test-Path $src)) {
    Write-Output ("ERROR: source not found: " + $src)
    exit 1
}

$img = [System.Drawing.Image]::FromFile($src)
$bmp = New-Object System.Drawing.Bitmap $img.Width, $img.Height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::Transparent)
$g.DrawImage($img, 0, 0, $img.Width, $img.Height)
$g.Dispose()
$img.Dispose()

# alpha bounding box: crop away the transparent margin so the artwork fills the icon
$minX = $bmp.Width; $minY = $bmp.Height; $maxX = -1; $maxY = -1
for ($y = 0; $y -lt $bmp.Height; $y++) {
    for ($x = 0; $x -lt $bmp.Width; $x++) {
        if ($bmp.GetPixel($x, $y).A -gt 20) {
            if ($x -lt $minX) { $minX = $x }
            if ($x -gt $maxX) { $maxX = $x }
            if ($y -lt $minY) { $minY = $y }
            if ($y -gt $maxY) { $maxY = $y }
        }
    }
}
if ($maxX -lt $minX -or $maxY -lt $minY) {
    $minX = 0; $minY = 0; $maxX = $bmp.Width - 1; $maxY = $bmp.Height - 1
}
$rect = New-Object System.Drawing.Rectangle $minX, $minY, ($maxX - $minX + 1), ($maxY - $minY + 1)
Write-Output ("content rect: " + $rect.Width + "x" + $rect.Height)

$crop = $bmp.Clone($rect, $bmp.PixelFormat)
$out = New-Object System.Drawing.Bitmap 64, 64, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g2 = [System.Drawing.Graphics]::FromImage($out)
$g2.Clear([System.Drawing.Color]::Transparent)
$g2.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g2.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g2.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$scale = [Math]::Min(64.0 / $rect.Width, 64.0 / $rect.Height)
$w = [int][Math]::Round($rect.Width * $scale)
$h = [int][Math]::Round($rect.Height * $scale)
$dx = [int]((64 - $w) / 2)
$dy = [int]((64 - $h) / 2)
$dest = New-Object System.Drawing.Rectangle $dx, $dy, $w, $h
$g2.DrawImage($crop, $dest)
$g2.Dispose()
$out.Save($dst, [System.Drawing.Imaging.ImageFormat]::Png)
$out.Dispose(); $crop.Dispose(); $bmp.Dispose()
Write-Output ("saved " + $dst + " (64x64)")
