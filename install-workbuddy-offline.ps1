param(
    [string]$InstallDir = "$HOME\.slidetwin\academic-tools-slidetwin",
    [switch]$RunPowerPointProbe,
    [switch]$SkipPythonPackages,
    [switch]$SkipMcpConfig
)

$ErrorActionPreference = "Stop"

function Refresh-ProcessPath {
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $segments = @(
        $machinePath,
        $userPath,
        "$env:ProgramFiles\nodejs",
        "$env:ProgramFiles\Git\cmd",
        "${env:ProgramFiles(x86)}\Git\cmd",
        "$env:LOCALAPPDATA\Programs\Python\Python313",
        "$env:LOCALAPPDATA\Programs\Python\Python313\Scripts",
        "$env:LOCALAPPDATA\Programs\Python\Python312",
        "$env:LOCALAPPDATA\Programs\Python\Python312\Scripts",
        "$env:LOCALAPPDATA\Programs\Python\Python311",
        "$env:LOCALAPPDATA\Programs\Python\Python311\Scripts",
        "$env:APPDATA\npm"
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $env:Path = $segments -join ";"
}

function Find-CommandPath {
    param([string[]]$Names, [string[]]$KnownPaths)

    Refresh-ProcessPath
    foreach ($name in $Names) {
        $command = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
            return $command.Source
        }
    }
    foreach ($candidate in $KnownPaths) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    return $null
}

function Invoke-Python {
    param([string]$PythonPath, [string[]]$Arguments)

    if ([IO.Path]::GetFileNameWithoutExtension($PythonPath) -ieq "py") {
        & $PythonPath -3 @Arguments
    }
    else {
        & $PythonPath @Arguments
    }
}

function Write-Utf8NoBom {
    param([string]$PathValue, [string]$Content)

    $parent = Split-Path -Parent $PathValue
    [System.IO.Directory]::CreateDirectory($parent) | Out-Null
    [System.IO.File]::WriteAllText(
        $PathValue,
        $Content,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

function Configure-WorkBuddyMcp {
    param(
        [string]$ConfigPath,
        [string]$NodePath,
        [string]$LauncherPath,
        [string]$RuntimeRoot
    )

    $config = $null
    if (Test-Path -LiteralPath $ConfigPath) {
        $backup = "$ConfigPath.backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Copy-Item -LiteralPath $ConfigPath -Destination $backup -Force
        Write-Host "Backed up existing MCP configuration: $backup"
        $text = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $config = $text | ConvertFrom-Json
        }
    }

    if ($null -eq $config) {
        $config = [pscustomobject]@{}
    }
    if (-not ($config.PSObject.Properties.Name -contains "mcpServers")) {
        $config | Add-Member -NotePropertyName "mcpServers" -NotePropertyValue ([pscustomobject]@{})
    }
    if ($null -eq $config.mcpServers) {
        $config.mcpServers = [pscustomobject]@{}
    }

    $server = [pscustomobject]@{
        type = "stdio"
        command = $NodePath
        args = @($LauncherPath)
        env = [pscustomobject]@{
            SLIDETWIN_RUNTIME_ROOT = $RuntimeRoot
        }
        description = "SlideTwin local editable-PowerPoint reconstruction tools"
    }

    $config.mcpServers | Add-Member -NotePropertyName "slidetwin-tools" -NotePropertyValue $server -Force
    Write-Utf8NoBom -PathValue $ConfigPath -Content ($config | ConvertTo-Json -Depth 20)
    Write-Host "Configured WorkBuddy MCP: $ConfigPath"
}

if ($env:OS -ne "Windows_NT") {
    throw "SlideTwin requires 64-bit Windows with desktop PowerPoint installed."
}

$payloadRoot = Join-Path $PSScriptRoot "payload"
if (-not (Test-Path -LiteralPath (Join-Path $payloadRoot "plugins\ppt-visual-reconstructor\scripts\mcp-server.mjs"))) {
    throw "Offline bundle payload is incomplete. Extract the ZIP before running this script."
}

$node = Find-CommandPath `
    -Names @("node.exe", "node") `
    -KnownPaths @("$env:ProgramFiles\nodejs\node.exe", "$env:LOCALAPPDATA\Programs\nodejs\node.exe")
if ($null -eq $node) {
    throw "Node.js was not found. Install Node.js LTS, reopen this installer, and run it again."
}

$python = Find-CommandPath `
    -Names @("py.exe", "py", "python.exe", "python") `
    -KnownPaths @(
        "$env:WINDIR\py.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe"
    )
if ($null -eq $python) {
    throw "Python 3.10 or newer was not found. Install Python, reopen this installer, and run it again."
}

Write-Host "SlideTwin offline installer for WorkBuddy"
Write-Host "Node.js: $node"
Write-Host "Python: $python"
Write-Host "Install directory: $InstallDir"

if (Test-Path -LiteralPath $InstallDir) {
    $backupDir = "$InstallDir.backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Move-Item -LiteralPath $InstallDir -Destination $backupDir
    Write-Host "Moved the previous incomplete installation to: $backupDir"
}

[System.IO.Directory]::CreateDirectory($InstallDir) | Out-Null
Copy-Item -Path (Join-Path $payloadRoot "*") -Destination $InstallDir -Recurse -Force

if (-not $SkipPythonPackages) {
    $requirements = Join-Path $InstallDir "requirements.txt"
    Write-Host "Installing Python image-processing dependencies..."
    Invoke-Python -PythonPath $python -Arguments @("-m", "pip", "install", "--user", "-r", $requirements) | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Python package installation failed."
    }
}

$runtimeRoot = Join-Path $InstallDir "plugins\ppt-visual-reconstructor"
$launcherPath = Join-Path $InstallDir "plugins\slidetwin-workbuddy\scripts\launch-mcp.mjs"

[Environment]::SetEnvironmentVariable("SLIDETWIN_RUNTIME_ROOT", $runtimeRoot, "User")
$env:SLIDETWIN_RUNTIME_ROOT = $runtimeRoot

if (-not $SkipMcpConfig) {
    $mcpConfigPath = Join-Path $HOME ".workbuddy\mcp.json"
    Configure-WorkBuddyMcp `
        -ConfigPath $mcpConfigPath `
        -NodePath $node `
        -LauncherPath $launcherPath `
        -RuntimeRoot $runtimeRoot
}

if ($RunPowerPointProbe) {
    $probe = Join-Path $runtimeRoot "scripts\probe_powerpoint_v4.ps1"
    $probeOutput = Join-Path $InstallDir "workbuddy-probe-output"
    Write-Host "Running the desktop PowerPoint bridge probe..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $probe -OutputDirectory $probeOutput | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "The runtime was installed, but the PowerPoint bridge probe failed."
    }
}

Write-Host ""
Write-Host "Runtime and MCP setup complete."
Write-Host "Restart WorkBuddy and check that slidetwin-tools appears under Connectors/MCP."
Write-Host "The original Codex plugin was not changed."
