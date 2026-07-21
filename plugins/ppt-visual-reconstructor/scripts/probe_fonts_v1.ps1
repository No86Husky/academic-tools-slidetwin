param(
    [string]$OutputDirectory = ""
)

$ErrorActionPreference = "Stop"
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
}
catch {
    # Continue when the host does not expose a configurable console encoding.
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $env:TEMP ("ppt-visual-reconstructor-fonts-" + [guid]::NewGuid().ToString("N"))
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$resultPath = Join-Path $OutputDirectory "font-probe.json"

$targets = @(
    "Noto Sans CJK SC",
    "Noto Sans SC",
    "Source Han Sans SC",
    "Source Han Sans CN",
    "Microsoft YaHei",
    "Microsoft YaHei UI"
)

try {
    Add-Type -AssemblyName System.Drawing
    $collection = New-Object System.Drawing.Text.InstalledFontCollection
    $installed = @($collection.Families | ForEach-Object { $_.Name } | Sort-Object -Unique)
    $targetStatus = @()
    foreach ($target in $targets) {
        $targetStatus += [ordered]@{
            name = $target
            installed = [bool]($installed -contains $target)
        }
    }
    $related = @(
        $installed | Where-Object {
            $_ -match "Noto" -or
            $_ -match "Source Han" -or
            $_ -match "YaHei"
        }
    )
    $result = [ordered]@{
        success = $true
        target_fonts = $targetStatus
        related_installed_fonts = $related
        installed_font_count = $installed.Count
    }
}
catch {
    $result = [ordered]@{
        success = $false
        error = $_.Exception.Message
        error_type = $_.Exception.GetType().FullName
    }
}

$result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resultPath -Encoding UTF8
$result | ConvertTo-Json -Depth 10
Write-Host "Font probe files: $OutputDirectory"

if (-not $result.success) {
    exit 1
}
