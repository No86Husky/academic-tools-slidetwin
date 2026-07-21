param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [Parameter(Mandatory = $true)]
    [string]$PlanPath,

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

$sourcePath = (Resolve-Path -LiteralPath $PresentationPath).Path
$resolvedPlanPath = (Resolve-Path -LiteralPath $PlanPath).Path
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $env:TEMP ("ppt-visual-reconstructor-apply-" + [guid]::NewGuid().ToString("N"))
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$renderDirectory = Join-Path $OutputDirectory "renders"
New-Item -ItemType Directory -Path $renderDirectory -Force | Out-Null
$outputPath = Join-Path $OutputDirectory "reconstructed.pptx"
$resultPath = Join-Path $OutputDirectory "application-result.json"

$plan = Get-Content -LiteralPath $resolvedPlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$protectedAddresses = @{}
foreach ($address in @($plan.protected_addresses)) {
    $protectedAddresses[[string]$address] = $true
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

function Resolve-ShapeAddress {
    param(
        [object]$Presentation,
        [string]$Address
    )
    $parts = @($Address -split "/")
    if ($parts.Count -lt 2) {
        throw "Invalid shape address: $Address"
    }
    $slideIndex = [int]$parts[0]
    $shapeIndex = [int]$parts[1]
    $shape = $Presentation.Slides.Item($slideIndex).Shapes.Item($shapeIndex)
    for ($index = 2; $index -lt $parts.Count; $index++) {
        if ([int]$shape.Type -ne 6) {
            throw "Address traverses a non-group shape: $Address"
        }
        $shape = $shape.GroupItems.Item([int]$parts[$index])
    }
    return $shape
}

function Get-ShapeFingerprint {
    param(
        [object]$Presentation,
        [string]$Address
    )
    $shape = Resolve-ShapeAddress -Presentation $Presentation -Address $Address
    return [ordered]@{
        address = $Address
        id = [int](Get-SafeProperty { $shape.Id } 0)
        name = [string](Get-SafeProperty { $shape.Name } "")
        type = [int](Get-SafeProperty { $shape.Type } -1)
        left_pt = [math]::Round([double](Get-SafeProperty { $shape.Left } 0), 4)
        top_pt = [math]::Round([double](Get-SafeProperty { $shape.Top } 0), 4)
        width_pt = [math]::Round([double](Get-SafeProperty { $shape.Width } 0), 4)
        height_pt = [math]::Round([double](Get-SafeProperty { $shape.Height } 0), 4)
        rotation_deg = [math]::Round([double](Get-SafeProperty { $shape.Rotation } 0), 4)
        crop_left_pt = [math]::Round([double](Get-SafeProperty { $shape.PictureFormat.CropLeft } 0), 4)
        crop_top_pt = [math]::Round([double](Get-SafeProperty { $shape.PictureFormat.CropTop } 0), 4)
        crop_right_pt = [math]::Round([double](Get-SafeProperty { $shape.PictureFormat.CropRight } 0), 4)
        crop_bottom_pt = [math]::Round([double](Get-SafeProperty { $shape.PictureFormat.CropBottom } 0), 4)
    }
}

function Get-TextShapeEntries {
    param(
        [object]$Shape,
        [string]$Address
    )
    $entries = @()
    $hasTextFrame = [int](Get-SafeProperty { $Shape.HasTextFrame } 0)
    if ($hasTextFrame -eq -1) {
        $hasText = [int](Get-SafeProperty { $Shape.TextFrame.HasText } 0)
        if ($hasText -eq -1) {
            $entries += [pscustomobject]@{
                Address = $Address
                Shape = $Shape
            }
        }
    }
    if ([int](Get-SafeProperty { $Shape.Type } -1) -eq 6) {
        $count = [int](Get-SafeProperty { $Shape.GroupItems.Count } 0)
        for ($childIndex = 1; $childIndex -le $count; $childIndex++) {
            $child = $Shape.GroupItems.Item($childIndex)
            $entries += Get-TextShapeEntries -Shape $child -Address ("$Address/$childIndex")
        }
    }
    return $entries
}

function Get-SlideTextShapeEntries {
    param(
        [object]$Presentation,
        [int]$SlideNumber
    )
    $entries = @()
    $slide = $Presentation.Slides.Item($SlideNumber)
    for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
        $shape = $slide.Shapes.Item($shapeIndex)
        $entries += Get-TextShapeEntries -Shape $shape -Address ("$SlideNumber/$shapeIndex")
    }
    return $entries
}

function Set-ScaledFont {
    param(
        [object]$Shape,
        [double]$Factor,
        [string]$Address
    )
    $textRange = $Shape.TextFrame.TextRange
    $before = [double]$textRange.Font.Size
    if ($before -le 0) {
        throw "Cannot scale a mixed or invalid font size at $Address"
    }
    $after = [math]::Round($before * $Factor, 3)
    $textRange.Font.Size = [single]$after
    return [ordered]@{
        op = "scale_font"
        address = $Address
        text = [string]$textRange.Text
        before_font_size_pt = $before
        factor = $Factor
        after_font_size_pt = [double]$textRange.Font.Size
    }
}

function Set-TextFontName {
    param(
        [object]$Shape,
        [string]$FontName,
        [string]$Address
    )
    if ([string]::IsNullOrWhiteSpace($FontName)) {
        throw "Font name cannot be empty at $Address"
    }
    $textRange = $Shape.TextFrame.TextRange
    $before = [string](Get-SafeProperty { $textRange.Font.Name } "")
    $textRange.Font.Name = $FontName
    $resolved = [string](Get-SafeProperty { $textRange.Font.Name } "")
    return [ordered]@{
        op = "set_font_name"
        address = $Address
        text = [string]$textRange.Text
        requested_font_name = $FontName
        before_font_name = $before
        resolved_font_name = $resolved
    }
}

function Convert-HexToOfficeRgb {
    param([string]$HexColor)
    if ($HexColor -notmatch '^#?([0-9A-Fa-f]{6})$') {
        throw "Invalid RGB color: $HexColor"
    }
    $hex = $Matches[1]
    $red = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return [int]($red -bor ($green -shl 8) -bor ($blue -shl 16))
}

function Convert-OfficeRgbToHex {
    param([object]$RgbValue)
    if ($null -eq $RgbValue) {
        return $null
    }
    $value = [int64]$RgbValue
    $red = $value -band 0xFF
    $green = ($value -shr 8) -band 0xFF
    $blue = ($value -shr 16) -band 0xFF
    return "#{0:X2}{1:X2}{2:X2}" -f $red, $green, $blue
}

Copy-Item -LiteralPath $sourcePath -Destination $outputPath -Force
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash

$powerPoint = $null
$presentation = $null
$stage = "initializing"
$operationLog = @()
$protectedBefore = @()
$protectedAfter = @()

try {
    $stage = "starting PowerPoint COM"
    Write-Host "[1/7] Starting PowerPoint COM..."
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.DisplayAlerts = 1

    $stage = "opening copied presentation"
    Write-Host "[2/7] Opening the copied presentation..."
    $presentation = $powerPoint.Presentations.Open($outputPath, 0, 0, 0)

    $stage = "capturing protected objects"
    Write-Host "[3/7] Capturing protected image fingerprints..."
    foreach ($address in @($protectedAddresses.Keys | Sort-Object)) {
        $protectedBefore += Get-ShapeFingerprint -Presentation $presentation -Address $address
    }

    $stage = "applying operations"
    Write-Host "[4/7] Applying the modification plan..."
    foreach ($operation in @($plan.operations)) {
        $opName = [string]$operation.op
        if ($opName -eq "scale_all_text") {
            $slideNumber = [int]$operation.slide_number
            $factor = [double]$operation.factor
            $excluded = @{}
            foreach ($address in @($operation.exclude_addresses)) {
                $excluded[[string]$address] = $true
            }
            $entries = Get-SlideTextShapeEntries -Presentation $presentation -SlideNumber $slideNumber
            foreach ($entry in $entries) {
                if ($protectedAddresses.ContainsKey($entry.Address)) {
                    throw "A protected object was selected for text scaling: $($entry.Address)"
                }
                if (-not $excluded.ContainsKey($entry.Address)) {
                    $operationLog += Set-ScaledFont -Shape $entry.Shape -Factor $factor -Address $entry.Address
                }
            }
        }
        elseif ($opName -eq "set_all_text_font") {
            $slideNumber = [int]$operation.slide_number
            $fontName = [string]$operation.font_name
            $excluded = @{}
            foreach ($address in @($operation.exclude_addresses)) {
                $excluded[[string]$address] = $true
            }
            $entries = Get-SlideTextShapeEntries -Presentation $presentation -SlideNumber $slideNumber
            foreach ($entry in $entries) {
                if ($protectedAddresses.ContainsKey($entry.Address)) {
                    throw "A protected object was selected for font replacement: $($entry.Address)"
                }
                if (-not $excluded.ContainsKey($entry.Address)) {
                    $operationLog += Set-TextFontName -Shape $entry.Shape -FontName $fontName -Address $entry.Address
                }
            }
        }
        elseif ($opName -eq "scale_font") {
            $address = [string]$operation.address
            if ($protectedAddresses.ContainsKey($address)) {
                throw "The plan attempts to modify a protected object: $address"
            }
            $shape = Resolve-ShapeAddress -Presentation $presentation -Address $address
            $operationLog += Set-ScaledFont -Shape $shape -Factor ([double]$operation.factor) -Address $address
        }
        elseif ($opName -eq "set_font_size") {
            $address = [string]$operation.address
            if ($protectedAddresses.ContainsKey($address)) {
                throw "The plan attempts to modify a protected object: $address"
            }
            $shape = Resolve-ShapeAddress -Presentation $presentation -Address $address
            $textRange = $shape.TextFrame.TextRange
            $before = [double]$textRange.Font.Size
            $textRange.Font.Size = [single]([double]$operation.font_size_pt)
            $operationLog += [ordered]@{
                op = "set_font_size"
                address = $address
                text = [string]$textRange.Text
                before_font_size_pt = $before
                after_font_size_pt = [double]$textRange.Font.Size
            }
        }
        elseif ($opName -eq "set_font_name") {
            $address = [string]$operation.address
            if ($protectedAddresses.ContainsKey($address)) {
                throw "The plan attempts to modify a protected object: $address"
            }
            $shape = Resolve-ShapeAddress -Presentation $presentation -Address $address
            $operationLog += Set-TextFontName -Shape $shape -FontName ([string]$operation.font_name) -Address $address
        }
        elseif ($opName -eq "set_font_spacing") {
            $address = [string]$operation.address
            if ($protectedAddresses.ContainsKey($address)) {
                throw "The plan attempts to modify a protected object: $address"
            }
            $shape = Resolve-ShapeAddress -Presentation $presentation -Address $address
            $font = $shape.TextFrame2.TextRange.Font
            $before = [double](Get-SafeProperty { $font.Spacing } 0)
            $font.Spacing = [single]([double]$operation.spacing_pt)
            $operationLog += [ordered]@{
                op = "set_font_spacing"
                address = $address
                text = [string]$shape.TextFrame.TextRange.Text
                before_spacing_pt = $before
                after_spacing_pt = [double](Get-SafeProperty { $font.Spacing } 0)
            }
        }
        elseif ($opName -eq "set_font_color") {
            $address = [string]$operation.address
            if ($protectedAddresses.ContainsKey($address)) {
                throw "The plan attempts to modify a protected object: $address"
            }
            $shape = Resolve-ShapeAddress -Presentation $presentation -Address $address
            $fontColor = $shape.TextFrame.TextRange.Font.Color
            $beforeRgb = Get-SafeProperty { $fontColor.RGB }
            $requestedHex = [string]$operation.color
            $fontColor.RGB = Convert-HexToOfficeRgb -HexColor $requestedHex
            $operationLog += [ordered]@{
                op = "set_font_color"
                address = $address
                text = [string]$shape.TextFrame.TextRange.Text
                before_color = Convert-OfficeRgbToHex $beforeRgb
                requested_color = $requestedHex.ToUpperInvariant()
                after_color = Convert-OfficeRgbToHex (Get-SafeProperty { $fontColor.RGB })
            }
        }
        elseif ($opName -eq "move_shape") {
            $address = [string]$operation.address
            if ($protectedAddresses.ContainsKey($address)) {
                throw "The plan attempts to modify a protected object: $address"
            }
            $shape = Resolve-ShapeAddress -Presentation $presentation -Address $address
            $beforeLeft = [double]$shape.Left
            $beforeTop = [double]$shape.Top
            $deltaLeft = [double](Get-SafeProperty { $operation.delta_left_pt } 0)
            $deltaTop = [double](Get-SafeProperty { $operation.delta_top_pt } 0)
            $shape.Left = [single]($beforeLeft + $deltaLeft)
            $shape.Top = [single]($beforeTop + $deltaTop)
            $operationLog += [ordered]@{
                op = "move_shape"
                address = $address
                delta_left_pt = $deltaLeft
                delta_top_pt = $deltaTop
                before_left_pt = $beforeLeft
                before_top_pt = $beforeTop
                after_left_pt = [double]$shape.Left
                after_top_pt = [double]$shape.Top
            }
        }
        elseif ($opName -eq "set_shape_visible") {
            $address = [string]$operation.address
            if ($protectedAddresses.ContainsKey($address)) {
                throw "The plan attempts to modify a protected object: $address"
            }
            $shape = Resolve-ShapeAddress -Presentation $presentation -Address $address
            $before = [int](Get-SafeProperty { $shape.Visible } 0)
            $after = if ([bool]$operation.visible) { -1 } else { 0 }
            $shape.Visible = $after
            $operationLog += [ordered]@{
                op = "set_shape_visible"
                address = $address
                before_visible = $before
                after_visible = [int](Get-SafeProperty { $shape.Visible } 0)
            }
        }
        elseif ($opName -eq "add_rounded_rectangle") {
            $slideNumber = [int]$operation.slide_number
            $slide = $presentation.Slides.Item($slideNumber)
            $shape = $slide.Shapes.AddShape(
                5,
                [single]([double]$operation.left_pt),
                [single]([double]$operation.top_pt),
                [single]([double]$operation.width_pt),
                [single]([double]$operation.height_pt)
            )
            if (-not [string]::IsNullOrWhiteSpace([string]$operation.name)) {
                $shape.Name = [string]$operation.name
            }
            $shape.Fill.Visible = -1
            $shape.Fill.Solid()
            $shape.Fill.ForeColor.RGB = Convert-HexToOfficeRgb -HexColor ([string]$operation.fill_color)
            $shape.Fill.Transparency = [single]([double]$operation.fill_transparency)
            if ([bool]$operation.line_visible) {
                $shape.Line.Visible = -1
                $shape.Line.ForeColor.RGB = Convert-HexToOfficeRgb -HexColor ([string]$operation.line_color)
                $shape.Line.Transparency = [single]([double]$operation.line_transparency)
                $shape.Line.Weight = [single]([double]$operation.line_weight_pt)
            }
            else {
                $shape.Line.Visible = 0
            }
            $adjustment = Get-SafeProperty { $operation.radius_adjustment } $null
            if ($null -ne $adjustment) {
                try {
                    $shape.Adjustments.Item(1) = [single]([double]$adjustment)
                }
                catch {
                    throw "Could not set the rounded-corner adjustment for $($shape.Name): $($_.Exception.Message)"
                }
            }
            $targetZOrder = [int](Get-SafeProperty { $operation.z_order } 0)
            if ($targetZOrder -gt 0) {
                while ([int]$shape.ZOrderPosition -gt $targetZOrder) {
                    $shape.ZOrder(3)
                }
            }
            $operationLog += [ordered]@{
                op = "add_rounded_rectangle"
                slide_number = $slideNumber
                name = [string]$shape.Name
                id = [int]$shape.Id
                left_pt = [double]$shape.Left
                top_pt = [double]$shape.Top
                width_pt = [double]$shape.Width
                height_pt = [double]$shape.Height
                radius_adjustment = [double](Get-SafeProperty { $shape.Adjustments.Item(1) } 0)
                z_order = [int]$shape.ZOrderPosition
            }
        }
        else {
            throw "Unsupported operation: $opName"
        }
    }

    $stage = "verifying protected objects"
    Write-Host "[5/7] Verifying protected images..."
    foreach ($address in @($protectedAddresses.Keys | Sort-Object)) {
        $protectedAfter += Get-ShapeFingerprint -Presentation $presentation -Address $address
    }
    $beforeJson = $protectedBefore | ConvertTo-Json -Depth 5 -Compress
    $afterJson = $protectedAfter | ConvertTo-Json -Depth 5 -Compress
    if ($beforeJson -ne $afterJson) {
        throw "A protected image fingerprint changed during modification."
    }

    $stage = "saving presentation"
    Write-Host "[6/7] Saving the reconstructed presentation..."
    $presentation.Save()

    $stage = "rendering slides"
    Write-Host "[7/7] Rendering slides with PowerPoint..."
    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight
    $renderHeight = [int][math]::Round($RenderWidth * $slideHeight / $slideWidth)
    $renders = @()
    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $renderPath = Join-Path $renderDirectory ("slide-{0}.png" -f $slideIndex)
        $presentation.Slides.Item($slideIndex).Export($renderPath, "PNG", $RenderWidth, $renderHeight)
        $renders += $renderPath
    }

    $result = [ordered]@{
        success = $true
        source_path = $sourcePath
        source_sha256 = $sourceHash
        plan_path = $resolvedPlanPath
        output_path = $outputPath
        operation_count = $operationLog.Count
        operations = $operationLog
        protected_objects_unchanged = $true
        protected_objects = $protectedAfter
        renders = $renders
    }
}
catch {
    $result = [ordered]@{
        success = $false
        stage = $stage
        source_path = $sourcePath
        source_sha256 = $sourceHash
        plan_path = $resolvedPlanPath
        output_path = $outputPath
        operations_completed = $operationLog.Count
        error = $_.Exception.Message
        error_type = $_.Exception.GetType().FullName
    }
}
finally {
    if ($presentation -ne $null) {
        try {
            $presentation.Saved = -1
            $presentation.Close()
        }
        catch {
            Write-Warning "Could not close the output presentation automatically."
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

if ($result.success) {
    $result.output_sha256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash
}
$result | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $resultPath -Encoding UTF8
$result | ConvertTo-Json -Depth 5
Write-Host "Application files: $OutputDirectory"

if (-not $result.success) {
    exit 1
}
