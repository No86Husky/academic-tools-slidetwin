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
    # Continue when the host does not expose a configurable console encoding.
}

$resolvedPresentation = (Resolve-Path -LiteralPath $PresentationPath).Path
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $env:TEMP ("ppt-visual-reconstructor-inspect-" + [guid]::NewGuid().ToString("N"))
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$renderDirectory = Join-Path $OutputDirectory "renders"
New-Item -ItemType Directory -Path $renderDirectory -Force | Out-Null
$resultPath = Join-Path $OutputDirectory "powerpoint-inspection.json"

$shapeTypeNames = @{
    1 = "auto_shape"
    5 = "freeform"
    6 = "group"
    9 = "line"
    11 = "linked_picture"
    13 = "picture"
    14 = "placeholder"
    17 = "text_box"
    19 = "table"
    24 = "smart_art"
    28 = "graphic"
}

function Convert-RgbToHex {
    param([object]$RgbValue)
    if ($null -eq $RgbValue) {
        return $null
    }
    try {
        $value = [long]$RgbValue
        $red = $value -band 0xFF
        $green = ($value -shr 8) -band 0xFF
        $blue = ($value -shr 16) -band 0xFF
        return ("#{0:X2}{1:X2}{2:X2}" -f $red, $green, $blue)
    }
    catch {
        return $null
    }
}

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

function Get-ShapeRecord {
    param(
        [object]$Shape,
        [string]$Address
    )

    $typeNumber = [int](Get-SafeProperty { $Shape.Type } -1)
    $typeName = if ($shapeTypeNames.ContainsKey($typeNumber)) {
        $shapeTypeNames[$typeNumber]
    }
    else {
        "type_$typeNumber"
    }

    $text = $null
    $textStyle = $null
    $hasTextFrame = [int](Get-SafeProperty { $Shape.HasTextFrame } 0)
    if ($hasTextFrame -eq -1) {
        $hasText = [int](Get-SafeProperty { $Shape.TextFrame.HasText } 0)
        if ($hasText -eq -1) {
            $textRange = $Shape.TextFrame.TextRange
            $text = [string](Get-SafeProperty { $textRange.Text } "")
            $textStyle = [ordered]@{
                font_name = [string](Get-SafeProperty { $textRange.Font.Name } "")
                font_size_pt = [double](Get-SafeProperty { $textRange.Font.Size } 0)
                bold = [int](Get-SafeProperty { $textRange.Font.Bold } 0)
                italic = [int](Get-SafeProperty { $textRange.Font.Italic } 0)
                color = Convert-RgbToHex (Get-SafeProperty { $textRange.Font.Color.RGB })
                alignment = [int](Get-SafeProperty { $textRange.ParagraphFormat.Alignment } 0)
                margin_left_pt = [double](Get-SafeProperty { $Shape.TextFrame.MarginLeft } 0)
                margin_right_pt = [double](Get-SafeProperty { $Shape.TextFrame.MarginRight } 0)
                margin_top_pt = [double](Get-SafeProperty { $Shape.TextFrame.MarginTop } 0)
                margin_bottom_pt = [double](Get-SafeProperty { $Shape.TextFrame.MarginBottom } 0)
                word_wrap = [int](Get-SafeProperty { $Shape.TextFrame.WordWrap } 0)
                auto_size = [int](Get-SafeProperty { $Shape.TextFrame.AutoSize } 0)
            }
        }
    }

    $fill = [ordered]@{
        visible = [int](Get-SafeProperty { $Shape.Fill.Visible } 0)
        color = Convert-RgbToHex (Get-SafeProperty { $Shape.Fill.ForeColor.RGB })
        transparency = [double](Get-SafeProperty { $Shape.Fill.Transparency } 0)
        type = [int](Get-SafeProperty { $Shape.Fill.Type } 0)
    }
    $line = [ordered]@{
        visible = [int](Get-SafeProperty { $Shape.Line.Visible } 0)
        color = Convert-RgbToHex (Get-SafeProperty { $Shape.Line.ForeColor.RGB })
        transparency = [double](Get-SafeProperty { $Shape.Line.Transparency } 0)
        weight_pt = [double](Get-SafeProperty { $Shape.Line.Weight } 0)
    }

    $picture = $null
    if ($typeNumber -eq 11 -or $typeNumber -eq 13 -or $typeNumber -eq 28) {
        $picture = [ordered]@{
            crop_left_pt = [double](Get-SafeProperty { $Shape.PictureFormat.CropLeft } 0)
            crop_right_pt = [double](Get-SafeProperty { $Shape.PictureFormat.CropRight } 0)
            crop_top_pt = [double](Get-SafeProperty { $Shape.PictureFormat.CropTop } 0)
            crop_bottom_pt = [double](Get-SafeProperty { $Shape.PictureFormat.CropBottom } 0)
            transparency_color = Convert-RgbToHex (Get-SafeProperty { $Shape.PictureFormat.TransparencyColor })
        }
    }

    $children = @()
    if ($typeNumber -eq 6) {
        $groupCount = [int](Get-SafeProperty { $Shape.GroupItems.Count } 0)
        for ($childIndex = 1; $childIndex -le $groupCount; $childIndex++) {
            $child = $Shape.GroupItems.Item($childIndex)
            $children += Get-ShapeRecord -Shape $child -Address ("$Address/$childIndex")
        }
    }

    return [ordered]@{
        address = $Address
        id = [int](Get-SafeProperty { $Shape.Id } 0)
        name = [string](Get-SafeProperty { $Shape.Name } "")
        type = $typeName
        type_number = $typeNumber
        left_pt = [math]::Round([double](Get-SafeProperty { $Shape.Left } 0), 3)
        top_pt = [math]::Round([double](Get-SafeProperty { $Shape.Top } 0), 3)
        width_pt = [math]::Round([double](Get-SafeProperty { $Shape.Width } 0), 3)
        height_pt = [math]::Round([double](Get-SafeProperty { $Shape.Height } 0), 3)
        rotation_deg = [math]::Round([double](Get-SafeProperty { $Shape.Rotation } 0), 3)
        z_order = [int](Get-SafeProperty { $Shape.ZOrderPosition } 0)
        visible = [int](Get-SafeProperty { $Shape.Visible } 0)
        locked_aspect_ratio = [int](Get-SafeProperty { $Shape.LockAspectRatio } 0)
        alternative_text = [string](Get-SafeProperty { $Shape.AlternativeText } "")
        text = $text
        text_style = $textStyle
        fill = $fill
        line = $line
        picture = $picture
        children = $children
    }
}

$powerPoint = $null
$presentation = $null
$stage = "initializing"

try {
    $stage = "starting PowerPoint COM"
    Write-Host "[1/5] Starting PowerPoint COM..."
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.DisplayAlerts = 1

    $stage = "opening presentation"
    Write-Host "[2/5] Opening the presentation read-only and without a window..."
    $presentation = $powerPoint.Presentations.Open($resolvedPresentation, -1, 0, 0)

    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight
    $renderHeight = [int][math]::Round($RenderWidth * $slideHeight / $slideWidth)
    $slides = @()

    $stage = "reading objects"
    Write-Host "[3/5] Reading slide objects and groups..."
    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        $objects = @()
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            $objects += Get-ShapeRecord -Shape $shape -Address ("$slideIndex/$shapeIndex")
        }

        $stage = "exporting slide $slideIndex"
        Write-Host ("[4/5] Exporting slide {0}/{1}..." -f $slideIndex, $presentation.Slides.Count)
        $renderPath = Join-Path $renderDirectory ("slide-{0}.png" -f $slideIndex)
        $slide.Export($renderPath, "PNG", $RenderWidth, $renderHeight)

        $slides += [ordered]@{
            slide_number = $slideIndex
            slide_id = [int]$slide.SlideID
            name = [string]$slide.Name
            top_level_object_count = [int]$slide.Shapes.Count
            render_path = $renderPath
            objects = $objects
        }
    }

    $stage = "collecting results"
    Write-Host "[5/5] Writing the inspection report..."
    $result = [ordered]@{
        success = $true
        source_path = $resolvedPresentation
        source_bytes = (Get-Item -LiteralPath $resolvedPresentation).Length
        powerpoint_version = [string]$powerPoint.Version
        slide_width_pt = [math]::Round($slideWidth, 3)
        slide_height_pt = [math]::Round($slideHeight, 3)
        render_width_px = $RenderWidth
        render_height_px = $renderHeight
        slide_count = [int]$presentation.Slides.Count
        slides = $slides
    }
}
catch {
    $result = [ordered]@{
        success = $false
        stage = $stage
        source_path = $resolvedPresentation
        error = $_.Exception.Message
        error_type = $_.Exception.GetType().FullName
    }
}
finally {
    if ($presentation -ne $null) {
        try {
            $presentation.Close()
        }
        catch {
            Write-Warning "Could not close the inspected presentation automatically."
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

$result | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $resultPath -Encoding UTF8
$result | ConvertTo-Json -Depth 5
Write-Host "Inspection files: $OutputDirectory"

if (-not $result.success) {
    exit 1
}
