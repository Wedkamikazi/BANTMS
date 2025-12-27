#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive Claude Code Installation Script for Windows 11

.DESCRIPTION
    This script installs the latest version of Claude Code CLI on Windows 11.
    It supports both the native Windows installation (recommended) and the NPM-based installation.

    As of December 2025, Claude Code supports native Windows installation without requiring
    WSL (Windows Subsystem for Linux) or Node.js.

.PARAMETER InstallMethod
    The installation method to use:
    - Native (default): Uses the official PowerShell installer (recommended)
    - NPM: Uses npm to install globally (requires Node.js 18+)

.PARAMETER Version
    The version to install:
    - stable (default): Latest stable release
    - latest: Most recent version including pre-releases
    - Specific version number (e.g., "1.0.58")

.EXAMPLE
    .\Install-ClaudeCode.ps1
    Installs the latest stable version using native installer.

.EXAMPLE
    .\Install-ClaudeCode.ps1 -InstallMethod NPM
    Installs using npm (requires Node.js).

.EXAMPLE
    .\Install-ClaudeCode.ps1 -Version latest
    Installs the latest version (including pre-releases).

.NOTES
    Author: BANTMS Project
    Date: December 2025
    Requirements: Windows 10 20H2+ or Windows 11

.LINK
    https://code.claude.com/docs/en/setup
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("Native", "NPM")]
    [string]$InstallMethod = "Native",

    [Parameter(Mandatory = $false)]
    [string]$Version = "stable"
)

#region Configuration
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# Colors for output
$Colors = @{
    Success = "Green"
    Warning = "Yellow"
    Error   = "Red"
    Info    = "Cyan"
    Header  = "Magenta"
}

# Installation paths
$NativeInstallPath = "$env:LOCALAPPDATA\Programs\claude-code"
$WindowsAppsPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps"
#endregion

#region Helper Functions
function Write-Header {
    param([string]$Text)
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor $Colors.Header
    Write-Host "  $Text" -ForegroundColor $Colors.Header
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor $Colors.Header
    Write-Host ""
}

function Write-Step {
    param(
        [int]$Number,
        [string]$Text
    )
    Write-Host "[$Number] $Text" -ForegroundColor $Colors.Info
}

function Write-Success {
    param([string]$Text)
    Write-Host "✓ $Text" -ForegroundColor $Colors.Success
}

function Write-Warning {
    param([string]$Text)
    Write-Host "⚠ $Text" -ForegroundColor $Colors.Warning
}

function Write-ErrorMessage {
    param([string]$Text)
    Write-Host "✗ $Text" -ForegroundColor $Colors.Error
}

function Write-Info {
    param([string]$Text)
    Write-Host "→ $Text" -ForegroundColor $Colors.Info
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-WindowsVersion {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    return @{
        Version     = $os.Version
        BuildNumber = $os.BuildNumber
        Caption     = $os.Caption
    }
}

function Test-CommandExists {
    param([string]$Command)
    $null -ne (Get-Command $Command -ErrorAction SilentlyContinue)
}

function Get-NodeVersion {
    if (Test-CommandExists "node") {
        $version = (node --version 2>$null)
        if ($version -match "v(\d+)\.(\d+)\.(\d+)") {
            return @{
                Major = [int]$Matches[1]
                Minor = [int]$Matches[2]
                Patch = [int]$Matches[3]
                Full  = $version
            }
        }
    }
    return $null
}

function Get-NpmVersion {
    if (Test-CommandExists "npm") {
        $version = (npm --version 2>$null)
        if ($version -match "(\d+)\.(\d+)\.(\d+)") {
            return @{
                Major = [int]$Matches[1]
                Minor = [int]$Matches[2]
                Patch = [int]$Matches[3]
                Full  = $version
            }
        }
    }
    return $null
}

function Test-ClaudeCodeInstalled {
    if (Test-CommandExists "claude") {
        try {
            $version = (claude --version 2>$null)
            return @{
                Installed = $true
                Version   = $version
            }
        }
        catch {
            return @{ Installed = $false }
        }
    }
    return @{ Installed = $false }
}
#endregion

#region System Checks
function Test-SystemRequirements {
    Write-Header "Checking System Requirements"

    $allPassed = $true

    # Check Windows version
    Write-Step 1 "Checking Windows version..."
    $winInfo = Get-WindowsVersion
    $buildNumber = [int]$winInfo.BuildNumber

    if ($buildNumber -ge 19041) {
        Write-Success "Windows version: $($winInfo.Caption) (Build $buildNumber)"
    }
    else {
        Write-ErrorMessage "Windows 10 Build 19041+ or Windows 11 required. Current: Build $buildNumber"
        $allPassed = $false
    }

    # Check PowerShell version
    Write-Step 2 "Checking PowerShell version..."
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -ge 5) {
        Write-Success "PowerShell version: $($psVersion.ToString())"
    }
    else {
        Write-ErrorMessage "PowerShell 5.1 or higher required. Current: $($psVersion.ToString())"
        $allPassed = $false
    }

    # Check internet connectivity
    Write-Step 3 "Checking internet connectivity..."
    try {
        $null = Invoke-WebRequest -Uri "https://claude.ai" -UseBasicParsing -TimeoutSec 10 -Method Head
        Write-Success "Internet connection available"
    }
    catch {
        Write-ErrorMessage "Cannot connect to claude.ai. Please check your internet connection."
        $allPassed = $false
    }

    # Check disk space
    Write-Step 4 "Checking available disk space..."
    $drive = (Get-Item $env:LOCALAPPDATA).PSDrive
    $freeSpaceGB = [math]::Round(($drive.Free / 1GB), 2)
    if ($freeSpaceGB -ge 1) {
        Write-Success "Available disk space: ${freeSpaceGB} GB"
    }
    else {
        Write-Warning "Low disk space: ${freeSpaceGB} GB. Recommended: 4 GB+"
    }

    # Check for existing installation
    Write-Step 5 "Checking for existing Claude Code installation..."
    $existing = Test-ClaudeCodeInstalled
    if ($existing.Installed) {
        Write-Warning "Claude Code is already installed: $($existing.Version)"
        Write-Info "This script will update/reinstall Claude Code."
    }
    else {
        Write-Success "No existing installation found"
    }

    # NPM-specific checks
    if ($InstallMethod -eq "NPM") {
        Write-Step 6 "Checking Node.js (required for NPM installation)..."
        $nodeVersion = Get-NodeVersion
        if ($nodeVersion) {
            if ($nodeVersion.Major -ge 18) {
                Write-Success "Node.js version: $($nodeVersion.Full)"
            }
            else {
                Write-ErrorMessage "Node.js 18+ required for NPM installation. Current: $($nodeVersion.Full)"
                $allPassed = $false
            }
        }
        else {
            Write-ErrorMessage "Node.js not found. Please install Node.js 18+ or use Native installation method."
            $allPassed = $false
        }

        Write-Step 7 "Checking npm..."
        $npmVersion = Get-NpmVersion
        if ($npmVersion) {
            Write-Success "npm version: $($npmVersion.Full)"
        }
        else {
            Write-ErrorMessage "npm not found."
            $allPassed = $false
        }
    }

    # Administrator check (not required but informative)
    Write-Step ($(if ($InstallMethod -eq "NPM") { 8 } else { 6 })) "Checking privileges..."
    if (Test-Administrator) {
        Write-Warning "Running as Administrator (not required for Claude Code installation)"
    }
    else {
        Write-Success "Running as standard user (recommended)"
    }

    Write-Host ""
    return $allPassed
}
#endregion

#region Installation Functions
function Install-ClaudeCodeNative {
    param([string]$VersionToInstall)

    Write-Header "Installing Claude Code (Native Method)"

    Write-Step 1 "Downloading and running official installer..."
    Write-Info "Version: $VersionToInstall"

    try {
        if ($VersionToInstall -eq "stable") {
            # Default stable installation
            Write-Info "Executing: irm https://claude.ai/install.ps1 | iex"
            $script = Invoke-RestMethod -Uri "https://claude.ai/install.ps1"
            Invoke-Expression $script
        }
        else {
            # Specific version or "latest"
            Write-Info "Executing installer with version: $VersionToInstall"
            $script = Invoke-RestMethod -Uri "https://claude.ai/install.ps1"
            $scriptBlock = [scriptblock]::Create($script)
            & $scriptBlock $VersionToInstall
        }

        Write-Success "Installation command completed"
        return $true
    }
    catch {
        Write-ErrorMessage "Installation failed: $($_.Exception.Message)"
        return $false
    }
}

function Install-ClaudeCodeNPM {
    Write-Header "Installing Claude Code (NPM Method)"

    Write-Step 1 "Installing @anthropic-ai/claude-code globally..."
    Write-Warning "Note: Do NOT use 'sudo' or run as Administrator for npm global installs"

    try {
        Write-Info "Executing: npm install -g @anthropic-ai/claude-code"
        $result = & npm install -g @anthropic-ai/claude-code 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Success "NPM installation completed"
            Write-Host $result
            return $true
        }
        else {
            Write-ErrorMessage "NPM installation failed"
            Write-Host $result -ForegroundColor $Colors.Error
            return $false
        }
    }
    catch {
        Write-ErrorMessage "Installation failed: $($_.Exception.Message)"
        return $false
    }
}

function Test-Installation {
    Write-Header "Verifying Installation"

    Write-Step 1 "Checking if 'claude' command is available..."

    # Refresh environment PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

    # Add common installation paths
    $pathsToCheck = @(
        $WindowsAppsPath,
        $NativeInstallPath,
        "$env:APPDATA\npm"
    )

    foreach ($path in $pathsToCheck) {
        if (Test-Path $path) {
            if ($env:Path -notlike "*$path*") {
                $env:Path = "$path;$env:Path"
            }
        }
    }

    Start-Sleep -Seconds 2

    if (Test-CommandExists "claude") {
        try {
            $version = & claude --version 2>&1
            Write-Success "Claude Code installed successfully!"
            Write-Host ""
            Write-Host "  Version: $version" -ForegroundColor $Colors.Success

            # Check installation location
            $claudePath = (Get-Command claude).Source
            Write-Host "  Location: $claudePath" -ForegroundColor $Colors.Info

            return $true
        }
        catch {
            Write-Warning "Claude command found but version check failed"
            return $true
        }
    }
    else {
        Write-ErrorMessage "Claude Code command not found in PATH"
        Write-Info "You may need to restart your terminal or add the installation directory to PATH"
        Write-Info "Expected locations:"
        Write-Info "  Native: $WindowsAppsPath\claude.exe"
        Write-Info "  NPM: $env:APPDATA\npm\claude.cmd"
        return $false
    }
}
#endregion

#region Post-Installation
function Show-NextSteps {
    Write-Header "Next Steps - Authentication Setup"

    Write-Host "To complete the setup, follow these steps:" -ForegroundColor $Colors.Info
    Write-Host ""

    Write-Host "  1. Open a NEW terminal window (to refresh PATH)" -ForegroundColor White
    Write-Host "  2. Navigate to your project directory:" -ForegroundColor White
    Write-Host "     cd C:\path\to\your\project" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  3. Start Claude Code:" -ForegroundColor White
    Write-Host "     claude" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  4. Choose your authentication method:" -ForegroundColor White
    Write-Host ""
    Write-Host "     Option A - Claude Console (Default):" -ForegroundColor $Colors.Info
    Write-Host "       • Select 'Claude Console' when prompted" -ForegroundColor Gray
    Write-Host "       • Complete OAuth in your browser" -ForegroundColor Gray
    Write-Host "       • Requires active billing at console.anthropic.com" -ForegroundColor Gray
    Write-Host ""
    Write-Host "     Option B - Claude App (Pro/Max Plan):" -ForegroundColor $Colors.Info
    Write-Host "       • Subscribe to Claude Pro or Max plan" -ForegroundColor Gray
    Write-Host "       • Unified billing for Claude Code + web interface" -ForegroundColor Gray
    Write-Host ""

    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor $Colors.Header
    Write-Host ""
    Write-Host "  Useful Commands:" -ForegroundColor $Colors.Info
    Write-Host "    claude --version    Check installed version" -ForegroundColor Gray
    Write-Host "    claude --help       Show help and available commands" -ForegroundColor Gray
    Write-Host "    claude config       Configure settings" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Documentation:" -ForegroundColor $Colors.Info
    Write-Host "    https://code.claude.com/docs" -ForegroundColor Gray
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor $Colors.Header
}

function Show-Uninstall {
    Write-Host ""
    Write-Host "  To uninstall Claude Code:" -ForegroundColor $Colors.Warning
    Write-Host ""

    if ($InstallMethod -eq "Native") {
        Write-Host "    # Remove native installation:" -ForegroundColor Gray
        Write-Host '    Remove-Item -Path "$env:LOCALAPPDATA\Programs\claude-code" -Recurse -Force' -ForegroundColor Gray
        Write-Host '    Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\WindowsApps\claude.exe" -Force' -ForegroundColor Gray
    }
    else {
        Write-Host "    # Remove NPM installation:" -ForegroundColor Gray
        Write-Host "    npm uninstall -g @anthropic-ai/claude-code" -ForegroundColor Gray
    }
    Write-Host ""
}
#endregion

#region Main Execution
function Main {
    Clear-Host

    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor $Colors.Header
    Write-Host "║                                                               ║" -ForegroundColor $Colors.Header
    Write-Host "║         CLAUDE CODE INSTALLER FOR WINDOWS 11                  ║" -ForegroundColor $Colors.Header
    Write-Host "║                   December 2025                               ║" -ForegroundColor $Colors.Header
    Write-Host "║                                                               ║" -ForegroundColor $Colors.Header
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor $Colors.Header
    Write-Host ""
    Write-Host "  Installation Method: $InstallMethod" -ForegroundColor $Colors.Info
    Write-Host "  Version: $Version" -ForegroundColor $Colors.Info
    Write-Host ""

    # Run system checks
    $checksPass = Test-SystemRequirements

    if (-not $checksPass) {
        Write-Host ""
        Write-ErrorMessage "System requirements not met. Please resolve the issues above and try again."
        Write-Host ""
        exit 1
    }

    # Confirm installation
    Write-Host ""
    $confirm = Read-Host "Proceed with installation? (Y/n)"
    if ($confirm -eq "n" -or $confirm -eq "N") {
        Write-Info "Installation cancelled by user."
        exit 0
    }

    # Run installation
    $installSuccess = $false

    switch ($InstallMethod) {
        "Native" {
            $installSuccess = Install-ClaudeCodeNative -VersionToInstall $Version
        }
        "NPM" {
            $installSuccess = Install-ClaudeCodeNPM
        }
    }

    if (-not $installSuccess) {
        Write-Host ""
        Write-ErrorMessage "Installation failed. Please check the error messages above."
        Write-Host ""
        exit 1
    }

    # Verify installation
    $verified = Test-Installation

    # Show next steps regardless of verification
    # (sometimes PATH updates require terminal restart)
    Show-NextSteps
    Show-Uninstall

    if ($verified) {
        Write-Success "Claude Code installation completed successfully!"
    }
    else {
        Write-Warning "Installation completed but verification pending."
        Write-Info "Please restart your terminal and run 'claude --version' to verify."
    }

    Write-Host ""
}

# Run the main function
Main
#endregion
