# Copyright (c) 2026 AbeLin · MIT License
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [int]$CropX = 8,
    [int]$CropY = 70,
    [int]$CropSize = 890
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$destination = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../assets'))
[System.IO.Directory]::CreateDirectory($destination) | Out-Null
$photo = [System.Drawing.Image]::FromFile((Resolve-Path -LiteralPath $Source).Path)

function New-IconFrame([int]$Size) {
    $bitmap = [System.Drawing.Bitmap]::new($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.DrawImage($photo, [System.Drawing.Rectangle]::new(0, 0, $Size, $Size), $CropX, $CropY, $CropSize, $CropSize, [System.Drawing.GraphicsUnit]::Pixel)
        return $bitmap
    } catch {
        $bitmap.Dispose()
        throw
    } finally { $graphics.Dispose() }
}

try {
    if ($CropX -lt 0 -or $CropY -lt 0 -or $CropSize -lt 1 -or ($CropX + $CropSize) -gt $photo.Width -or ($CropY + $CropSize) -gt $photo.Height) {
        throw '裁切范围超出原图。'
    }
    $preview = New-IconFrame 512
    try { $preview.Save((Join-Path $destination 'app.png'), [System.Drawing.Imaging.ImageFormat]::Png) }
    finally { $preview.Dispose() }

    # PNG 压缩的多帧 ICO，同一原图分别采样，适配任务栏、高 DPI 和资源管理器。
    $sizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)
    $frames = [System.Collections.Generic.List[byte[]]]::new()
    foreach ($size in $sizes) {
        $frame = New-IconFrame $size
        $memory = [System.IO.MemoryStream]::new()
        try {
            $frame.Save($memory, [System.Drawing.Imaging.ImageFormat]::Png)
            $frames.Add($memory.ToArray())
        } finally { $memory.Dispose(); $frame.Dispose() }
    }
    $stream = [System.IO.File]::Create((Join-Path $destination 'app.ico'))
    $writer = [System.IO.BinaryWriter]::new($stream)
    try {
        $writer.Write([uint16]0)
        $writer.Write([uint16]1)
        $writer.Write([uint16]$sizes.Count)
        $offset = 6 + 16 * $sizes.Count
        for ($index = 0; $index -lt $sizes.Count; $index++) {
            $dimension = if ($sizes[$index] -eq 256) { 0 } else { $sizes[$index] }
            $writer.Write([byte]$dimension); $writer.Write([byte]$dimension)
            $writer.Write([byte]0); $writer.Write([byte]0)
            $writer.Write([uint16]1); $writer.Write([uint16]32)
            $writer.Write([uint32]$frames[$index].Length); $writer.Write([uint32]$offset)
            $offset += $frames[$index].Length
        }
        foreach ($frameBytes in $frames) { $writer.Write($frameBytes) }
    } finally { $writer.Dispose(); $stream.Dispose() }
    [PSCustomObject]@{ Png = (Join-Path $destination 'app.png'); Icon = (Join-Path $destination 'app.ico'); Sizes = $sizes } | ConvertTo-Json
} finally { $photo.Dispose() }
