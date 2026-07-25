param(
    [string]$InstallDir = "$HOME\.slidetwin\academic-tools-slidetwin",
    [string]$RepositoryRef = "feat/workbuddy-plugin",
    [switch]$InstallPrerequisites,
    [switch]$RunPowerPointProbe,
    [switch]$SkipPythonPackages,
    [switch]$SkipMcpConfig,
    [switch]$DoNotOpenSkillPackage
)

$ErrorActionPreference = "Stop"
$Repository = "https://github.com/No86Husky/academic-tools-slidetwin.git"

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
}
catch {
}

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
    ) | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_)
    }
    $env:Path = $segments -join ";"
}

function Find-CommandPath {
    param([string[]]$Names)

    foreach ($name in $Names) {
        $command = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
            return $command.Source
        }
    }
    return $null
}

function Find-KnownPath {
    param([string[]]$Candidates)

    foreach ($candidate in $Candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    return $null
}

function Resolve-Or-InstallTool {
    param(
        [string[]]$Names,
        [string[]]$KnownPaths,
        [string]$WingetId,
        [string]$Label
    )

    Refresh-ProcessPath

    $path = Find-CommandPath -Names $Names
    if ($null -eq $path) {
        $path = Find-KnownPath -Candidates $KnownPaths
    }
    if ($null -ne $path) {
        Write-Host "Found ${Label}: $path"
        return $path
    }

    if (-not $InstallPrerequisites) {
        throw "$Label is required but was not found. Run the installer again with -InstallPrerequisites."
    }

    $winget = Find-CommandPath -Names @("winget.exe", "winget")
    if ($null -eq $winget) {
        throw "winget is required for automatic installation of $Label. Install $Label manually, reopen PowerShell, and rerun this installer."
    }

    Write-Host "Installing or verifying $Label ($WingetId)..."
    & $winget install --id $WingetId -e --source winget --accept-package-agreements --accept-source-agreements | Out-Host
    $wingetExitCode = $LASTEXITCODE
    if ($wingetExitCode -ne 0) {
        Write-Host "winget returned exit code $wingetExitCode. This can mean the package is already installed and has no available update."
    }

    Refresh-ProcessPath
    $path = Find-CommandPath -Names $Names
    if ($null -eq $path) {
        $path = Find-KnownPath -Candidates $KnownPaths
    }

    if ($null -eq $path) {
        throw "$Label is still not available after the winget attempt. Close this window, reopen PowerShell, confirm that $Label is installed, and rerun the installer."
    }

    Write-Host "Found $Label after verification: $path"
    return $path
}

function Write-Utf8NoBom {
    param(
        [string]$PathValue,
        [string]$Content
    )

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
            try {
                $config = $text | ConvertFrom-Json
            }
            catch {
                throw "Existing WorkBuddy MCP configuration is not valid JSON: $ConfigPath"
            }
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
    Write-Host "Configured WorkBuddy user-level MCP: $ConfigPath"
}

function Invoke-Python {
    param(
        [string]$PythonPath,
        [string[]]$Arguments
    )

    $leaf = [IO.Path]::GetFileNameWithoutExtension($PythonPath)
    if ($leaf -ieq "py") {
        & $PythonPath -3 @Arguments
    }
    else {
        & $PythonPath @Arguments
    }
}

if ($env:OS -ne "Windows_NT") {
    throw "The editable PowerPoint workflow requires 64-bit Windows with desktop PowerPoint installed."
}

Write-Host "SlideTwin installer for WorkBuddy desktop"
Write-Host "Repository ref: $RepositoryRef"
Write-Host "Install directory: $InstallDir"
Write-Host "This installer does not install CodeBuddy Code and does not change the Codex plugin."

$git = Resolve-Or-InstallTool `
    -Names @("git.exe", "git") `
    -KnownPaths @(
        "$env:ProgramFiles\Git\cmd\git.exe",
        "${env:ProgramFiles(x86)}\Git\cmd\git.exe"
    ) `
    -WingetId "Git.Git" `
    -Label "Git"

$node = Resolve-Or-InstallTool `
    -Names @("node.exe", "node") `
    -KnownPaths @(
        "$env:ProgramFiles\nodejs\node.exe",
        "$env:LOCALAPPDATA\Programs\nodejs\node.exe"
    ) `
    -WingetId "OpenJS.NodeJS.LTS" `
    -Label "Node.js LTS"

$python = Resolve-Or-InstallTool `
    -Names @("py.exe", "py", "python.exe", "python") `
    -KnownPaths @(
        "$env:WINDIR\py.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe"
    ) `
    -WingetId "Python.Python.3.12" `
    -Label "Python 3.12"

$parent = Split-Path -Parent $InstallDir
New-Item -ItemType Directory -Path $parent -Force | Out-Null

if (Test-Path -LiteralPath (Join-Path $InstallDir ".git")) {
    Write-Host "Updating the existing SlideTwin checkout..."
    & $git -C $InstallDir fetch origin $RepositoryRef | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Could not fetch repository ref $RepositoryRef."
    }

    & $git -C $InstallDir checkout $RepositoryRef | Out-Host
    if ($LASTEXITCODE -ne 0) {
        & $git -C $InstallDir checkout -B $RepositoryRef "origin/$RepositoryRef" | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw "Could not check out repository ref $RepositoryRef."
        }
    }

    & $git -C $InstallDir pull --ff-only origin $RepositoryRef | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Could not update the checkout. Resolve local changes in $InstallDir and rerun the installer."
    }
}
elseif (Test-Path -LiteralPath $InstallDir) {
    $existingItems = @(Get-ChildItem -LiteralPath $InstallDir -Force -ErrorAction SilentlyContinue)
    if ($existingItems.Count -eq 0) {
        Remove-Item -LiteralPath $InstallDir -Force
        Write-Host "Downloading SlideTwin..."
        & $git clone --branch $RepositoryRef --single-branch $Repository $InstallDir | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw "Could not clone $Repository at ref $RepositoryRef."
        }
    }
    else {
        throw "Install directory already exists but is not a Git checkout: $InstallDir"
    }
}
else {
    Write-Host "Downloading SlideTwin..."
    & $git clone --branch $RepositoryRef --single-branch $Repository $InstallDir | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Could not clone $Repository at ref $RepositoryRef."
    }
}

if (-not $SkipPythonPackages) {
    Write-Host "Installing Python image-processing dependencies..."
    $requirements = Join-Path $InstallDir "requirements.txt"
    Invoke-Python -PythonPath $python -Arguments @("-m", "pip", "install", "--user", "-r", $requirements) | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Python package installation failed."
    }
}

$runtimeRoot = Join-Path $InstallDir "plugins\ppt-visual-reconstructor"
$launcherPath = Join-Path $InstallDir "plugins\slidetwin-workbuddy\scripts\launch-mcp.mjs"

if (-not (Test-Path -LiteralPath (Join-Path $runtimeRoot "scripts\mcp-server.mjs"))) {
    throw "The shared SlideTwin runtime is incomplete: $runtimeRoot"
}

if (-not (Test-Path -LiteralPath $launcherPath)) {
    throw "The WorkBuddy MCP launcher is missing: $launcherPath"
}

[Environment]::SetEnvironmentVariable("SLIDETWIN_RUNTIME_ROOT", $runtimeRoot, "User")
$env:SLIDETWIN_RUNTIME_ROOT = $runtimeRoot
Write-Host "SlideTwin runtime: $runtimeRoot"

if (-not $SkipMcpConfig) {
    $mcpConfigPath = Join-Path $HOME ".workbuddy\mcp.json"
    Configure-WorkBuddyMcp `
        -ConfigPath $mcpConfigPath `
        -NodePath $node `
        -LauncherPath $launcherPath `
        -RuntimeRoot $runtimeRoot
}

$skillOutput = Join-Path $InstallDir "dist\slidetwin-workbuddy-skill.zip"
$skillBuilder = Join-Path $InstallDir "tools\build_workbuddy_skill.py"

Write-Host "Building the WorkBuddy Skill package..."
Invoke-Python -PythonPath $python -Arguments @($skillBuilder, "--output", $skillOutput) | Out-Host
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $skillOutput)) {
    throw "Could not build the WorkBuddy Skill package."
}

if ($RunPowerPointProbe) {
    $probe = Join-Path $runtimeRoot "scripts\probe_powerpoint_v4.ps1"
    $probeOutput = Join-Path $InstallDir "workbuddy-probe-output"

    Write-Host "Running the desktop PowerPoint bridge probe..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $probe -OutputDirectory $probeOutput | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "The SlideTwin runtime is configured, but the PowerPoint bridge probe failed. Review the output above."
    }
}

Write-Host ""
Write-Host "Runtime and MCP setup complete."
Write-Host "Skill package: $skillOutput"
Write-Host ""
Write-Host "Restart WorkBuddy, open Connectors/MCP, and confirm that slidetwin-tools is connected."
Write-Host "The original Codex plugin was not changed, installed, updated, or removed."

if (-not $DoNotOpenSkillPackage) {
    try {
        Start-Process explorer.exe -ArgumentList "/select,`"$skillOutput`""
    }
    catch {
        Write-Host "Could not open File Explorer automatically. Open the Skill package path manually."
    }
}
