param(
    [string]$InstallDir = "$HOME\.codex\marketplaces\ppt-visual-tools",
    [switch]$InstallPrerequisites,
    [switch]$RunPowerPointProbe,
    [switch]$SkipPythonPackages
)

$ErrorActionPreference = "Stop"
$Repository = "https://github.com/No86Husky/academic-tools.git"
$Marketplace = "ppt-visual-tools"
$Plugin = "ppt-visual-reconstructor@$Marketplace"

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
        throw "$Label was installed but is not visible in this PowerShell process. Close PowerShell, open it again, and rerun the installer."
    }
    return $path
}

Write-Host "SlideTwin installer"
Write-Host "Install directory: $InstallDir"

$git = Require-Or-Install -Names @("git.exe", "git") -WingetId "Git.Git" -Label "Git"
$node = Require-Or-Install -Names @("node.exe", "node") -WingetId "OpenJS.NodeJS.LTS" -Label "Node.js LTS"
$python = Require-Or-Install -Names @("py.exe", "py", "python.exe", "python") -WingetId "Python.Python.3.12" -Label "Python 3.12"

$codex = Find-Command @("codex.cmd", "codex.exe", "codex")
if ($null -eq $codex) {
    if (-not $InstallPrerequisites) {
        throw "Codex CLI is required but was not found. Run this installer again with -InstallPrerequisites."
    }
    $npm = Find-Command @("npm.cmd", "npm.exe", "npm")
    if ($null -eq $npm) {
        throw "npm was not found after Node.js installation. Close PowerShell, open it again, and rerun the installer."
    }
    Write-Host "Installing Codex CLI..."
    & $npm install --global @openai/codex
    if ($LASTEXITCODE -ne 0) {
        throw "npm could not install Codex CLI. Exit code: $LASTEXITCODE"
    }
    Refresh-ProcessPath
    $codex = Find-Command @("codex.cmd", "codex.exe", "codex")
    if ($null -eq $codex) {
        throw "Codex CLI was installed but is not visible in this PowerShell process. Close PowerShell, open it again, and rerun the installer."
    }
}

$parent = Split-Path -Parent $InstallDir
New-Item -ItemType Directory -Path $parent -Force | Out-Null

if (Test-Path -LiteralPath (Join-Path $InstallDir ".git")) {
    Write-Host "Updating the existing marketplace checkout..."
    & $git -C $InstallDir pull --ff-only
    if ($LASTEXITCODE -ne 0) {
        throw "Could not update the existing checkout. Resolve local Git changes in $InstallDir and rerun the installer."
    }
}
elseif (Test-Path -LiteralPath $InstallDir) {
    throw "Install directory already exists but is not a Git checkout: $InstallDir"
}
else {
    Write-Host "Downloading the plugin repository..."
    & $git clone $Repository $InstallDir
    if ($LASTEXITCODE -ne 0) {
        throw "Could not clone $Repository"
    }
}

if (-not $SkipPythonPackages) {
    Write-Host "Installing Python image-comparison packages for the current user..."
    $pythonLeaf = [IO.Path]::GetFileNameWithoutExtension($python)
    if ($pythonLeaf -ieq "py") {
        & $python -3 -m pip install --user -r (Join-Path $InstallDir "requirements.txt")
    }
    else {
        & $python -m pip install --user -r (Join-Path $InstallDir "requirements.txt")
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Python package installation failed."
    }
}

Write-Host "Registering the Codex marketplace..."
& $codex plugin marketplace add $InstallDir
if ($LASTEXITCODE -ne 0) {
    throw "Codex could not register the marketplace."
}

Write-Host "Installing $Plugin..."
& $codex plugin add $Plugin
if ($LASTEXITCODE -ne 0) {
    throw "Codex could not install $Plugin."
}

if ($RunPowerPointProbe) {
    $probe = Join-Path $InstallDir "plugins\ppt-visual-reconstructor\scripts\probe_powerpoint_v4.ps1"
    Write-Host "Running the desktop PowerPoint bridge probe..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $probe
    if ($LASTEXITCODE -ne 0) {
        throw "The plugin is installed, but the PowerPoint bridge probe failed. Review the probe output above."
    }
}

Write-Host ""
Write-Host "Installation complete."
Write-Host "Close and reopen Codex, start a new thread, upload one slide image, and ask:"
Write-Host 'Use $slidetwin to recreate this image as an editable PowerPoint slide.'
Write-Host 'Compatibility: $ppt-visual-reconstructor remains supported for existing prompts and videos.'
