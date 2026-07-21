param(
    [string]$OutputDirectory = ""
)

$ErrorActionPreference = "Stop"

# Keep the script source ASCII-only so Windows PowerShell 5.1 cannot decode
# embedded CJK literals with the wrong legacy code page.
$probeText = "PowerPoint 2024 " + (-join ([char[]](0x4E2D, 0x6587, 0x6E32, 0x67D3, 0x6D4B, 0x8BD5)))
$fontName = -join ([char[]](0x5FAE, 0x8F6F, 0x96C5, 0x9ED1))
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
}
catch {
    # Continue when the host does not expose a configurable console encoding.
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $env:TEMP ("ppt-visual-reconstructor-probe-" + [guid]::NewGuid().ToString("N"))
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$pptxPath = Join-Path $OutputDirectory "powerpoint-probe.pptx"
$pngPath = Join-Path $OutputDirectory "powerpoint-probe.png"
$resultPath = Join-Path $OutputDirectory "powerpoint-probe.json"

$powerPoint = $null
$presentation = $null
$stage = "initializing"

try {
    $stage = "starting PowerPoint COM"
    Write-Host "[1/7] Starting PowerPoint COM..."
    $powerPoint = New-Object -ComObject PowerPoint.Application

    $stage = "creating hidden presentation"
    Write-Host "[2/7] Creating a hidden presentation..."
    $powerPoint.DisplayAlerts = 1
    # WithWindow=0 avoids the visible startup window and third-party add-in UI.
    $presentation = $powerPoint.Presentations.Add(0)

    $stage = "adding test slide"
    Write-Host "[3/7] Adding a test slide..."
    $slide = $presentation.Slides.Add(1, 12)

    $stage = "adding test objects"
    Write-Host "[4/7] Adding Chinese text and a shape..."
    $title = $slide.Shapes.AddTextbox(1, 72, 72, 720, 72)
    $title.Name = "probe_chinese_text"
    $title.TextFrame.TextRange.Text = $probeText
    $title.TextFrame.TextRange.Font.Name = $fontName
    $title.TextFrame.TextRange.Font.Size = 30

    $marker = $slide.Shapes.AddShape(1, 72, 180, 240, 80)
    $marker.Name = "probe_shape"
    $marker.Fill.ForeColor.RGB = 10731500
    $marker.Line.Visible = 0

    $stage = "saving PPTX"
    Write-Host "[5/7] Saving the PPTX..."
    $presentation.SaveAs($pptxPath, 24)

    $stage = "exporting PNG"
    Write-Host "[6/7] Exporting the slide as PNG..."
    $slide.Export($pngPath, "PNG", 1600, 900)

    $stage = "collecting results"
    Write-Host "[7/7] Collecting results..."
    $result = [ordered]@{
        success = $true
        windows = [System.Environment]::OSVersion.VersionString
        powershell_process_bits = ([IntPtr]::Size * 8)
        powerpoint_version = $powerPoint.Version
        pptx_path = $pptxPath
        png_path = $pngPath
        pptx_bytes = (Get-Item $pptxPath).Length
        png_bytes = (Get-Item $pngPath).Length
        chinese_text = $title.TextFrame.TextRange.Text
        requested_font = $fontName
        resolved_font = $title.TextFrame.TextRange.Font.Name
    }
}
catch {
    $result = [ordered]@{
        success = $false
        stage = $stage
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
            Write-Warning "Could not close the test presentation automatically: $($_.Exception.Message)"
        }
    }
    if ($powerPoint -ne $null) {
        try {
            $powerPoint.Quit()
        }
        catch {
            Write-Warning "Could not close PowerPoint automatically: $($_.Exception.Message)"
        }
    }
}

$result | ConvertTo-Json -Depth 5 | Set-Content -Path $resultPath -Encoding UTF8
$result | ConvertTo-Json -Depth 5
Write-Host "Probe files: $OutputDirectory"

if (-not $result.success) {
    exit 1
}
