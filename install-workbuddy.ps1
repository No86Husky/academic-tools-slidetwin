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
    $segments = @($machinePath, $userPath, "$env:APPDATA\npm") | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_)
    }
    $env:Path = $segments -join ";"
}

function Find-Command {
    param([string[]]$Names)
    foreach ($name in $Names) {
        $command = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command) {
            return $command.Source
        }
    }
    return $null
}

function Install-WingetPackage {
    param([string]$Id, [string]$Label)
    Write-Host "Installing $Label ($Id)..."
    & winget.exe install --id $Id -e --source winget --accept-package-agreements --accept-source-agreements | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "winget could not install $Label. Exit code: $LASTEXITCODE"
    }
    Refresh-ProcessPath
}

function Require-Or-Install {
    param(
        [string[]]$Names,
        [string]$WingetId,
        [string]$Label
    )
    $path = Find-Command $Names
    if ($null -ne $path) {
        return $path
    }
    if (-not $InstallPrerequisites) {
        throw "$Label is required but was not found. Run this installer again with -InstallPrerequisites."
    }
    if ($null -eq (Find-Command @("winget.exe", "winget"))) {
        throw "winget is required for automatic prerequisite installation. Install $Label manually, then rerun this script."
    }
    Install-WingetPackage -Id $WingetId -Label $Label
    $path = Find-Command $Names
    if ($null -eq $path) {
        throw "$Label was installed but is not visible in this PowerShell process. Reopen PowerShell and rerun the installer."
    }
    return $path
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

if ($env:OS -ne "Windows_NT") {
    throw "The editable PowerPoint workflow requires 64-bit Windows with desktop PowerPoint installed."
}

Write-Host "SlideTwin installer for WorkBuddy desktop"
Write-Host "Repository ref: $RepositoryRef"
Write-Host "Install directory: $InstallDir"
Write-Host "This installer does not install CodeBuddy Code and does not change the Codex plugin."

$git = Require-Or-Install -Names @("git.exe", "git") -WingetId "Git.Git" -Label "Git"
$node = Require-Or-Install -Names @("node.exe", "node") -WingetId "OpenJS.NodeJS.LTS" -Label "Node.js LTS"
$python = Require-Or-Install -Names @("py.exe", "py", "python.exe", "python") -WingetId "Python.Python.3.12" -Label "Python 3.12"

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
    }
    & $git -C $InstallDir pull --ff-only origin $RepositoryRef | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Could not update the checkout. Resolve local changes in $InstallDir and rerun the installer."
    }
}
elif (Test-Path -LiteralPath $InstallDir) {
    throw "Install directory already exists but is not a Git checkout: $InstallDir"
}
else {
    Write-Host "Downloading SlideTwin..."
    & $git clone --branch $RepositoryRef --single-branch $Repository $InstallDir | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Could not clone $Repository at ref $RepositoryRef"
    }
}

if (-not $SkipPythonPackages) {
    Write-Host "Installing Python image-processing dependencies..."
    $requirements = Join-Path $InstallDir "requirements.txt"
    $pythonLeaf = [IO.Path]::GetFileNameWithoutExtension($python)
    if ($pythonLeaf -ieq "py") {
        & $python -3 -m pip install --user -r $requirements | Out-Host
    }
    else {
        & $python -m pip install --user -r $requirements | Out-Host
    }
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
$pythonLeaf = [IO.Path]::GetFileNameWithoutExtension($python)
Write-Host "Building the WorkBuddy Skill package..."
if ($pythonLeaf -ieq "py") {
    & $python -3 $skillBuilder --output $skillOutput | Out-Host
}
else {
    & $python $skillBuilder --output $skillOutput | Out-Host
}
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
Write-Host "In WorkBuddy: Skills -> Add Skill -> Upload Skill, then select the ZIP above."
Write-Host "After upload, open Connectors/MCP and confirm slidetwin-tools is green."
Write-Host "The original Codex plugin was not changed, installed, updated, or removed."

if (-not $DoNotOpenSkillPackage) {
    try {
        Start-Process explorer.exe -ArgumentList "/select,`"$skillOutput`""
    }
    catch {
        Write-Host "Could not open File Explorer automatically. Open the Skill package path manually."
    }
}
