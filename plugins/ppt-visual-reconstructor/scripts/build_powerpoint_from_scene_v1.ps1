param(
    [Parameter(Mandatory = $true)]
    [string]$ReferenceImagePath,

    [Parameter(Mandatory = $true)]
    [string]$ScenePlanPath,

    [string]$OutputDirectory = "",

    [int]$RenderWidth = 1600
)

$ErrorActionPreference = "Stop"
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
}
catch {
}

function Write-Step {
    param([string]$Message)
    Write-Host $Message
}

function Resolve-ExistingFile {
    param([string]$PathValue, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        throw "$Label path is empty"
    }
    $resolved = [System.IO.Path]::GetFullPath($PathValue)
    if (-not [System.IO.File]::Exists($resolved)) {
        throw "$Label does not exist: $resolved"
    }
    return $resolved
}

function Convert-HexToOfficeRgb {
    param([string]$HexColor)
    if ([string]::IsNullOrWhiteSpace($HexColor)) {
        $HexColor = "#000000"
    }
    if ($HexColor -notmatch '^#[0-9A-Fa-f]{6}$') {
        throw "Color must use #RRGGBB: $HexColor"
    }
    $red = [Convert]::ToInt32($HexColor.Substring(1, 2), 16)
    $green = [Convert]::ToInt32($HexColor.Substring(3, 2), 16)
    $blue = [Convert]::ToInt32($HexColor.Substring(5, 2), 16)
    return ($red + ($green * 256) + ($blue * 65536))
}

function Get-NumberOrDefault {
    param($Value, [double]$Default)
    if ($null -eq $Value) {
        return $Default
    }
    return [double]$Value
}

function Get-BooleanOrDefault {
    param($Value, [bool]$Default)
    if ($null -eq $Value) {
        return $Default
    }
    return [bool]$Value
}

function Set-ShapeFill {
    param($Shape, $FillDefinition)
    if ($null -eq $FillDefinition -or [string]::IsNullOrWhiteSpace([string]$FillDefinition.color)) {
        $Shape.Fill.Visible = 0
        return
    }
    $Shape.Fill.Visible = -1
    $Shape.Fill.Solid()
    $Shape.Fill.ForeColor.RGB = Convert-HexToOfficeRgb ([string]$FillDefinition.color)
    $opacity = Get-NumberOrDefault $FillDefinition.opacity 1.0
    $Shape.Fill.Transparency = [Math]::Max(0.0, [Math]::Min(1.0, 1.0 - $opacity))
}

function Set-ShapeLine {
    param($Shape, $StrokeDefinition)
    if ($null -eq $StrokeDefinition -or [string]::IsNullOrWhiteSpace([string]$StrokeDefinition.color)) {
        $Shape.Line.Visible = 0
        return
    }
    $Shape.Line.Visible = -1
    $Shape.Line.ForeColor.RGB = Convert-HexToOfficeRgb ([string]$StrokeDefinition.color)
    $Shape.Line.Weight = Get-NumberOrDefault $StrokeDefinition.width_pt 1.0
    $opacity = Get-NumberOrDefault $StrokeDefinition.opacity 1.0
    $Shape.Line.Transparency = [Math]::Max(0.0, [Math]::Min(1.0, 1.0 - $opacity))
    if ($null -ne $StrokeDefinition.dash_style) {
        $dashMap = @{
            "solid" = 1
            "square_dot" = 2
            "round_dot" = 3
            "dash" = 4
            "dash_dot" = 5
        }
        $dashName = [string]$StrokeDefinition.dash_style
        if ($dashMap.ContainsKey($dashName)) {
            $Shape.Line.DashStyle = $dashMap[$dashName]
        }
    }
}

function Set-ShapeShadow {
    param($Shape, $ShadowDefinition)
    if ($null -eq $ShadowDefinition) {
        $Shape.Shadow.Visible = 0
        return
    }
    $Shape.Shadow.Visible = -1
    $Shape.Shadow.ForeColor.RGB = Convert-HexToOfficeRgb ([string]$ShadowDefinition.color)
    $opacity = Get-NumberOrDefault $ShadowDefinition.opacity 0.1
    $Shape.Shadow.Transparency = [Math]::Max(0.0, [Math]::Min(1.0, 1.0 - $opacity))
    $Shape.Shadow.OffsetX = Get-NumberOrDefault $ShadowDefinition.offset_x_pt 0
    $Shape.Shadow.OffsetY = Get-NumberOrDefault $ShadowDefinition.offset_y_pt 2
    try {
        $Shape.Shadow.Blur = Get-NumberOrDefault $ShadowDefinition.blur_pt 4
    }
    catch {
    }
}

function Get-BoundsInPoints {
    param($Bounds, [double]$ScaleX, [double]$ScaleY)
    if ($null -eq $Bounds -or $Bounds.Count -ne 4) {
        throw "bounds_px must contain four numbers"
    }
    return @(
        ([double]$Bounds[0] * $ScaleX),
        ([double]$Bounds[1] * $ScaleY),
        ([double]$Bounds[2] * $ScaleX),
        ([double]$Bounds[3] * $ScaleY)
    )
}

function Set-ElementIdentity {
    param($Shape, $Element)
    $targetName = [string]$Element.name
    if ([string]::IsNullOrWhiteSpace($targetName)) {
        $targetName = [string]$Element.id
    }
    if (-not [string]::IsNullOrWhiteSpace($targetName)) {
        $Shape.Name = $targetName
    }
    if ($null -ne $Element.rotation) {
        $Shape.Rotation = [double]$Element.rotation
    }
    try {
        $Shape.AlternativeText = "scene-id=" + [string]$Element.id + ";classification=" + [string]$Element.classification
    }
    catch {
    }
}

function Add-NativeShape {
    param($Slide, $Element, [double]$ScaleX, [double]$ScaleY)
    $shapeTypes = @{
        "rectangle" = 1
        "rounded_rectangle" = 5
        "ellipse" = 9
        "triangle" = 7
        "diamond" = 4
        "chevron" = 52
    }
    $shapeTypeName = [string]$Element.shape_type
    if (-not $shapeTypes.ContainsKey($shapeTypeName)) {
        throw "Unsupported shape_type: $shapeTypeName"
    }
    $bounds = Get-BoundsInPoints $Element.bounds_px $ScaleX $ScaleY
    $shape = $Slide.Shapes.AddShape($shapeTypes[$shapeTypeName], $bounds[0], $bounds[1], $bounds[2], $bounds[3])
    Set-ElementIdentity $shape $Element
    Set-ShapeFill $shape $Element.fill
    Set-ShapeLine $shape $Element.stroke
    Set-ShapeShadow $shape $Element.shadow
    return $shape
}

function Add-NativeLine {
    param($Slide, $Element, [double]$ScaleX, [double]$ScaleY)
    if ($null -eq $Element.points_px -or $Element.points_px.Count -ne 4) {
        throw "points_px must contain four numbers for line " + [string]$Element.id
    }
    $shape = $Slide.Shapes.AddLine(
        ([double]$Element.points_px[0] * $ScaleX),
        ([double]$Element.points_px[1] * $ScaleY),
        ([double]$Element.points_px[2] * $ScaleX),
        ([double]$Element.points_px[3] * $ScaleY)
    )
    Set-ElementIdentity $shape $Element
    Set-ShapeLine $shape $Element.stroke
    if ($null -ne $Element.stroke.begin_arrow) {
        $shape.Line.BeginArrowheadStyle = 3
    }
    if ($null -ne $Element.stroke.end_arrow) {
        $shape.Line.EndArrowheadStyle = 3
    }
    return $shape
}

function Add-NativeText {
    param($Slide, $Element, [double]$ScaleX, [double]$ScaleY)
    $bounds = Get-BoundsInPoints $Element.bounds_px $ScaleX $ScaleY
    $shape = $Slide.Shapes.AddTextbox(1, $bounds[0], $bounds[1], $bounds[2], $bounds[3])
    Set-ElementIdentity $shape $Element
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0

    $textFrame = $shape.TextFrame2
    $textFrame.AutoSize = 0
    $paragraph = $Element.paragraph
    $textFrame.WordWrap = if (Get-BooleanOrDefault $paragraph.word_wrap $true) { -1 } else { 0 }
    $textFrame.MarginLeft = Get-NumberOrDefault $paragraph.margin_left_pt 0
    $textFrame.MarginRight = Get-NumberOrDefault $paragraph.margin_right_pt 0
    $textFrame.MarginTop = Get-NumberOrDefault $paragraph.margin_top_pt 0
    $textFrame.MarginBottom = Get-NumberOrDefault $paragraph.margin_bottom_pt 0
    $verticalMap = @{
        "top" = 1
        "middle" = 3
        "bottom" = 4
    }
    $verticalName = [string]$paragraph.vertical_align
    if (-not $verticalMap.ContainsKey($verticalName)) {
        $verticalName = "top"
    }
    $textFrame.VerticalAnchor = $verticalMap[$verticalName]

    $range = $textFrame.TextRange
    $range.Text = [string]$Element.text
    $font = $Element.font
    $range.Font.Name = [string]$font.family
    $range.Font.NameFarEast = [string]$font.family
    $range.Font.Size = Get-NumberOrDefault $font.size_pt 18
    $range.Font.Bold = if (Get-BooleanOrDefault $font.bold $false) { -1 } else { 0 }
    $range.Font.Italic = if (Get-BooleanOrDefault $font.italic $false) { -1 } else { 0 }
    $range.Font.Fill.Visible = -1
    $range.Font.Fill.Solid()
    $range.Font.Fill.ForeColor.RGB = Convert-HexToOfficeRgb ([string]$font.color)
    if ($null -ne $font.spacing_pt) {
        $range.Font.Spacing = [double]$font.spacing_pt
    }
    $alignMap = @{
        "left" = 1
        "center" = 2
        "right" = 3
        "justify" = 4
    }
    $alignName = [string]$paragraph.align
    if (-not $alignMap.ContainsKey($alignName)) {
        $alignName = "left"
    }
    $range.ParagraphFormat.Alignment = $alignMap[$alignName]
    if ($null -ne $paragraph.space_before_pt) {
        $range.ParagraphFormat.SpaceBefore = [double]$paragraph.space_before_pt
    }
    if ($null -ne $paragraph.space_after_pt) {
        $range.ParagraphFormat.SpaceAfter = [double]$paragraph.space_after_pt
    }
    if ($null -ne $paragraph.line_spacing_pt) {
        $range.ParagraphFormat.SpaceWithin = [double]$paragraph.line_spacing_pt
    }
    return $shape
}

function Add-PictureElement {
    param($Slide, $Element, [double]$ScaleX, [double]$ScaleY)
    $assetPath = [string]$Element.source.asset_path
    if (-not [System.IO.File]::Exists($assetPath)) {
        throw "Resolved image asset is missing: $assetPath"
    }
    $bounds = Get-BoundsInPoints $Element.bounds_px $ScaleX $ScaleY
    $shape = $Slide.Shapes.AddPicture($assetPath, 0, -1, $bounds[0], $bounds[1], $bounds[2], $bounds[3])
    Set-ElementIdentity $shape $Element
    return $shape
}

function Release-ComObjectSafely {
    param($Object)
    if ($null -ne $Object) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
        }
        catch {
        }
    }
}

$powerPoint = $null
$presentation = $null
$slide = $null
$resultPath = $null

try {
    Write-Step "[1/8] Resolving the reference image and scene plan..."
    $referencePath = Resolve-ExistingFile $ReferenceImagePath "Reference image"
    $planPath = Resolve-ExistingFile $ScenePlanPath "Scene plan"
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $OutputDirectory = Join-Path ([System.IO.Path]::GetDirectoryName($planPath)) "image-only-output-$stamp"
    }
    $outputPath = [System.IO.Path]::GetFullPath($OutputDirectory)
    [System.IO.Directory]::CreateDirectory($outputPath) | Out-Null
    $resultPath = Join-Path $outputPath "image-only-build-result.json"

    Write-Step "[2/8] Preparing raster regions and semantic SVG..."
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($null -eq $python) {
        throw "Python was not found on PATH. Install Python 3.10 or newer and reopen PowerShell."
    }
    $prepareScript = Join-Path $PSScriptRoot "prepare_image_scene.py"
    & $python.Source $prepareScript --reference $referencePath --plan $planPath --output-dir $outputPath
    if ($LASTEXITCODE -ne 0) {
        throw "Scene preparation failed with exit code $LASTEXITCODE"
    }
    $resolvedPlanPath = Join-Path $outputPath "scene-plan.resolved.json"
    if (-not [System.IO.File]::Exists($resolvedPlanPath)) {
        throw "Resolved scene plan was not created"
    }
    $scene = Get-Content -Raw -Encoding UTF8 $resolvedPlanPath | ConvertFrom-Json
    $scenePreparationResultPath = Join-Path $outputPath "scene-preparation-result.json"
    $scenePreparation = Get-Content -Raw -Encoding UTF8 $scenePreparationResultPath | ConvertFrom-Json

    Write-Step "[3/8] Starting desktop PowerPoint..."
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1

    Write-Step "[4/8] Creating a blank presentation and native slide canvas..."
    $presentation = $powerPoint.Presentations.Add()
    $presentation.PageSetup.SlideWidth = [double]$scene.slide.width_pt
    $presentation.PageSetup.SlideHeight = [double]$scene.slide.height_pt
    $slide = $presentation.Slides.Add(1, 12)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Visible = -1
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = Convert-HexToOfficeRgb ([string]$scene.slide.background)

    $scaleX = [double]$scene.slide.width_pt / [double]$scene.slide.width_px
    $scaleY = [double]$scene.slide.height_pt / [double]$scene.slide.height_px
    $counts = @{
        native_text = 0
        native_shape = 0
        svg_object = 0
        raster_picture = 0
    }

    Write-Step "[5/8] Building native text, shapes, lines, and protected pictures..."
    $orderedElements = @($scene.elements | Sort-Object @{Expression = { [int]$_.z }}, @{Expression = { [string]$_.id }})
    foreach ($element in $orderedElements) {
        if ([bool]$element.exclude_from_final -or [bool]$element.guide_only) {
            continue
        }
        $created = $null
        switch ([string]$element.kind) {
            "shape" { $created = Add-NativeShape $slide $element $scaleX $scaleY }
            "line" { $created = Add-NativeLine $slide $element $scaleX $scaleY }
            "text" { $created = Add-NativeText $slide $element $scaleX $scaleY }
            "image" { $created = Add-PictureElement $slide $element $scaleX $scaleY }
            default { throw "Unsupported element kind: " + [string]$element.kind }
        }
        $classification = [string]$element.classification
        if ($counts.ContainsKey($classification)) {
            $counts[$classification] += 1
        }
        Release-ComObjectSafely $created
    }

    Write-Step "[6/8] Saving the editable PPTX..."
    $pptxPath = Join-Path $outputPath "reconstructed-from-image.pptx"
    $presentation.SaveAs($pptxPath, 24)

    Write-Step "[7/8] Exporting a native PowerPoint render..."
    $renderHeight = [int][Math]::Round($RenderWidth * [double]$scene.slide.height_px / [double]$scene.slide.width_px)
    $pngPath = Join-Path $outputPath "reconstructed-from-image.png"
    $slide.Export($pngPath, "PNG", $RenderWidth, $renderHeight)

    Write-Step "[8/8] Writing the build report..."
    $result = [ordered]@{
        success = $true
        reference_image = $referencePath
        source_scene_plan = $planPath
        resolved_scene_plan = $resolvedPlanPath
        semantic_svg = (Join-Path $outputPath "semantic-preview.svg")
        pptx_path = $pptxPath
        png_path = $pngPath
        slide_width_pt = [double]$scene.slide.width_pt
        slide_height_pt = [double]$scene.slide.height_pt
        element_count = [int]$orderedElements.Count
        native_text_count = [int]$counts.native_text
        native_shape_count = [int]$counts.native_shape
        svg_object_count = [int]$counts.svg_object
        protected_picture_count = [int]$counts.raster_picture
        editable_element_count = [int]$scenePreparation.editable_element_count
        editable_element_ratio = [double]$scenePreparation.editable_element_ratio
        editable_area_ratio_estimate = [double]$scenePreparation.editable_area_ratio_estimate
    }
    [System.IO.File]::WriteAllText($resultPath, ($result | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding($false)))
    $result | ConvertTo-Json -Depth 10
}
catch {
    $errorResult = [ordered]@{
        success = $false
        error = $_.Exception.Message
        error_type = $_.Exception.GetType().FullName
    }
    if ($null -ne $resultPath) {
        try {
            [System.IO.File]::WriteAllText($resultPath, ($errorResult | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding($false)))
        }
        catch {
        }
    }
    $errorResult | ConvertTo-Json -Depth 10
    exit 1
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
    }
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch {}
    }
    Release-ComObjectSafely $slide
    Release-ComObjectSafely $presentation
    Release-ComObjectSafely $powerPoint
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
