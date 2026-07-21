param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [string]$OutputDirectory = "",

    [int]$RenderWidth = 1600
)

$ErrorActionPreference = "Stop"
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
}
catch {
    # Continue when the host does not expose configurable output encoding.
}

$sourcePath = (Resolve-Path -LiteralPath $PresentationPath).Path
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $env:TEMP ("ppt-visual-reconstructor-font-sweep-" + [guid]::NewGuid().ToString("N"))
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$renderDirectory = Join-Path $OutputDirectory "renders"
New-Item -ItemType Directory -Path $renderDirectory -Force | Out-Null
$resultPath = Join-Path $OutputDirectory "font-sweep-result.json"
$zipPath = Join-Path $OutputDirectory "font-sweep-renders.zip"

$variants = @(
    [pscustomobject]@{ slug = "00-current"; font_name = $null },
    [pscustomobject]@{ slug = "01-microsoft-yahei-ui"; font_name = "Microsoft YaHei UI" },
    [pscustomobject]@{ slug = "02-microsoft-yahei-ui-light"; font_name = "Microsoft YaHei UI Light" },
    [pscustomobject]@{ slug = "03-noto-sans-sc"; font_name = "Noto Sans SC" },
    [pscustomobject]@{ slug = "04-noto-sans-sc-medium"; font_name = "Noto Sans SC Medium" },
    [pscustomobject]@{ slug = "05-noto-sans-sc-demilight"; font_name = "Noto Sans SC DemiLight" },
    [pscustomobject]@{ slug = "06-noto-sans-sc-light"; font_name = "Noto Sans SC Light" }
)

function Get-SafeProperty {
    param(
        [scriptblock]$Getter,
        [object]$DefaultValue = $null
    )
    try {
        return (& $Getter)
    }
    catch {
        return $DefaultValue
    }
}

function Get-TextShapes {
    param([object]$Shape)
    $items = @()
    $hasTextFrame = [int](Get-SafeProperty { $Shape.HasTextFrame } 0)
    if ($hasTextFrame -eq -1) {
        $hasText = [int](Get-SafeProperty { $Shape.TextFrame.HasText } 0)
        if ($hasText -eq -1) {
            $items += $Shape
        }
    }
    if ([int](Get-SafeProperty { $Shape.Type } -1) -eq 6) {
        $count = [int](Get-SafeProperty { $Shape.GroupItems.Count } 0)
        for ($index = 1; $index -le $count; $index++) {
            $items += Get-TextShapes -Shape $Shape.GroupItems.Item($index)
        }
    }
    return $items
}

$powerPoint = $null
$presentation = $null
$results = @()
$stage = "initializing"

try {
    $stage = "starting PowerPoint COM"
    Write-Host "[1/4] Starting PowerPoint COM..."
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.DisplayAlerts = 1

    $variantNumber = 0
    foreach ($variant in $variants) {
        $variantNumber++
        $stage = "rendering font variant $($variant.slug)"
        Write-Host ("[2/4] Rendering variant {0}/{1}: {2}" -f $variantNumber, $variants.Count, $variant.slug)
        $variantPath = Join-Path $OutputDirectory ($variant.slug + ".pptx")
        Copy-Item -LiteralPath $sourcePath -Destination $variantPath -Force
        $presentation = $powerPoint.Presentations.Open($variantPath, 0, 0, 0)

        $textShapeCount = 0
        $resolvedFonts = @{}
        for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
            $slide = $presentation.Slides.Item($slideIndex)
            for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
                $shape = $slide.Shapes.Item($shapeIndex)
                foreach ($textShape in @(Get-TextShapes -Shape $shape)) {
                    $textShapeCount++
                    if ($null -ne $variant.font_name) {
                        $textShape.TextFrame.TextRange.Font.Name = [string]$variant.font_name
                    }
                    $resolvedName = [string](Get-SafeProperty { $textShape.TextFrame.TextRange.Font.Name } "")
                    if (-not $resolvedFonts.ContainsKey($resolvedName)) {
                        $resolvedFonts[$resolvedName] = 0
                    }
                    $resolvedFonts[$resolvedName]++
                }
            }
        }

        $presentation.Save()
        $slideWidth = [double]$presentation.PageSetup.SlideWidth
        $slideHeight = [double]$presentation.PageSetup.SlideHeight
        $renderHeight = [int][math]::Round($RenderWidth * $slideHeight / $slideWidth)
        $renderPaths = @()
        for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
            $renderPath = Join-Path $renderDirectory (("{0}-slide-{1}.png" -f $variant.slug, $slideIndex))
            $presentation.Slides.Item($slideIndex).Export($renderPath, "PNG", $RenderWidth, $renderHeight)
            $renderPaths += $renderPath
        }

        $results += [ordered]@{
            slug = [string]$variant.slug
            requested_font = if ($null -eq $variant.font_name) { "(current fonts)" } else { [string]$variant.font_name }
            resolved_fonts = @($resolvedFonts.GetEnumerator() | Sort-Object Name | ForEach-Object {
                [ordered]@{ name = [string]$_.Name; text_shape_count = [int]$_.Value }
            })
            text_shape_count = $textShapeCount
            renders = $renderPaths
        }

        $presentation.Saved = -1
        $presentation.Close()
        $presentation = $null
        Remove-Item -LiteralPath $variantPath -Force
    }

    $stage = "writing results"
    Write-Host "[3/4] Writing font sweep results..."
    $result = [ordered]@{
        success = $true
        source_path = $sourcePath
        source_sha256 = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
        variants = $results
    }
    $result | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $resultPath -Encoding UTF8

    $stage = "creating render archive"
    Write-Host "[4/4] Creating render archive..."
    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    Compress-Archive -Path (Join-Path $renderDirectory "*.png") -DestinationPath $zipPath -CompressionLevel Optimal
}
catch {
    $result = [ordered]@{
        success = $false
        stage = $stage
        source_path = $sourcePath
        variants_completed = $results.Count
        error = $_.Exception.Message
        error_type = $_.Exception.GetType().FullName
    }
    $result | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $resultPath -Encoding UTF8
}
finally {
    if ($presentation -ne $null) {
        try {
            $presentation.Saved = -1
            $presentation.Close()
        }
        catch {
            Write-Warning "Could not close the active font-sweep presentation automatically."
        }
    }
    if ($powerPoint -ne $null) {
        try {
            $powerPoint.Quit()
        }
        catch {
            Write-Warning "Could not close PowerPoint automatically."
        }
    }
}

$result | ConvertTo-Json -Depth 8
Write-Host "Font sweep files: $OutputDirectory"

if (-not $result.success) {
    exit 1
}
