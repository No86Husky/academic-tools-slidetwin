param(
    [string]$InstallDir = "$HOME\.slidetwin\academic-tools-slidetwin",
    [string]$RepositoryRef = "feat/workbuddy-plugin",
    [string]$MarketplaceName = "slidetwin-tools",
    [switch]$InstallPrerequisites,
    [switch]$RunPowerPointProbe,
    [switch]$SkipPythonPackages
)

$ErrorActionPreference = "Stop"
$Repository = "https://github.com/No86Husky/academic-tools-slidetwin.git"
$Plugin = "slidetwin@$MarketplaceName"

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
        throw "$Label is required but was not found. Run the installer again with -InstallPrerequisites."
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

function Invoke-CodeBuddy {
    param([string[]]$Arguments, [switch]$AllowFailure)
    & $script:CodeBuddy @Arguments | Out-Host
    $exitCode = $LASTEXITCODE
    if (-not $AllowFailure -and $exitCode -ne 0) {
        throw "CodeBuddy command failed: codebuddy $($Arguments -join ' ') (exit code $exitCode)"
    }
    return $exitCode
}

if ($env:OS -ne "Windows_NT") {
    throw "The editable PowerPoint workflow requires 64-bit Windows with desktop PowerPoint installed."
}

Write-Host "SlideTwin WorkBuddy installer"
Write-Host "Repository ref: $RepositoryRef"
Write-Host "Install directory: $InstallDir"

$git = Require-Or-Install -Names @("git.exe", "git") -WingetId "Git.Git" -Label "Git"
$node = Require-Or-Install -Names @("node.exe", "node") -WingetId "OpenJS.NodeJS.LTS" -Label "Node.js LTS"
$python = Require-Or-Install -Names @("py.exe", "py", "python.exe", "python") -WingetId "Python.Python.3.12" -Label "Python 3.12"

$script:CodeBuddy = Find-Command @("codebuddy.cmd", "codebuddy.exe", "codebuddy")
if ($null -eq $script:CodeBuddy) {
    if (-not $InstallPrerequisites) {
        throw "CodeBuddy Code is required but was not found. Install it or rerun with -InstallPrerequisites."
    }
    $npm = Find-Command @("npm.cmd", "npm.exe", "npm")
    if ($null -eq $npm) {
        throw "npm was not found after Node.js installation. Reopen PowerShell and rerun the installer."
    }
    Write-Host "Installing CodeBuddy Code..."
    & $npm install --global @tencent-ai/codebuddy-code | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "npm could not install CodeBuddy Code. Exit code: $LASTEXITCODE"
    }
    Refresh-ProcessPath
    $script:CodeBuddy = Find-Command @("codebuddy.cmd", "codebuddy.exe", "codebuddy")
    if ($null -eq $script:CodeBuddy) {
        throw "CodeBuddy Code was installed but is not visible. Reopen PowerShell and rerun the installer."
    }
}

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
elseif (Test-Path -LiteralPath $InstallDir) {
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
if (-not (Test-Path -LiteralPath (Join-Path $runtimeRoot "scripts\mcp-server.mjs"))) {
    throw "The shared SlideTwin runtime is incomplete: $runtimeRoot"
}
[Environment]::SetEnvironmentVariable("SLIDETWIN_RUNTIME_ROOT", $runtimeRoot, "User")
$env:SLIDETWIN_RUNTIME_ROOT = $runtimeRoot
Write-Host "Shared runtime: $runtimeRoot"

Write-Host "Registering the local WorkBuddy marketplace..."
$marketplaceExit = Invoke-CodeBuddy -Arguments @("plugin", "marketplace", "add", $InstallDir, "--name", $MarketplaceName) -AllowFailure
if ($marketplaceExit -ne 0) {
    Write-Host "Marketplace may already exist; updating it instead..."
    Invoke-CodeBuddy -Arguments @("plugin", "marketplace", "update", $MarketplaceName)
}

Write-Host "Installing the WorkBuddy plugin..."
$installExit = Invoke-CodeBuddy -Arguments @("plugin", "install", $Plugin, "--scope", "user") -AllowFailure
if ($installExit -ne 0) {
    Write-Host "Plugin may already be installed; updating it instead..."
    Invoke-CodeBuddy -Arguments @("plugin", "update", $Plugin, "--scope", "user")
}

if ($RunPowerPointProbe) {
    $probe = Join-Path $runtimeRoot "scripts\probe_powerpoint_v4.ps1"
    $probeOutput = Join-Path $InstallDir "workbuddy-probe-output"
    Write-Host "Running the desktop PowerPoint bridge probe..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $probe -OutputDirectory $probeOutput | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "The WorkBuddy plugin is installed, but the PowerPoint bridge probe failed. Review the output above."
    }
}

Write-Host ""
Write-Host "Installation complete."
Write-Host "Start CodeBuddy Code, run /reload-plugins, upload one slide image, and use:"
Write-Host "/slidetwin:reconstruct"
Write-Host ""
Write-Host "The WorkBuddy plugin uses the shared SlideTwin runtime at:"
Write-Host $runtimeRoot
Write-Host "The original Codex plugin has not been changed or reinstalled."
