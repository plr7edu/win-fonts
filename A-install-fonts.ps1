#Requires -Version 5.1

<#
.SYNOPSIS
    Installs fonts system-wide using FontRegister.

.DESCRIPTION
    This script installs all .ttf and .otf font files from the current directory
    system-wide using FontRegister. Requires administrator privileges.

.NOTES
    Author: PowerShell Script
    Requires: Administrator rights
#>

# Configuration
$FontRegisterPath = "C:\#PL-DEV\1_DEV\Windows_Fonts\FontRegister-net48-win-x64\FontRegister.exe"

# Check if running as Administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Main Script
Clear-Host

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Font Installation Utility" -ForegroundColor White
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Check administrator rights
if (-not (Test-Administrator)) {
    Write-Host "ERROR: Administrator privileges required!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please restart this script as Administrator:" -ForegroundColor Yellow
    Write-Host "  1. Right-click on PowerShell" -ForegroundColor Gray
    Write-Host "  2. Select 'Run as Administrator'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Verify FontRegister exists
if (-not (Test-Path $FontRegisterPath)) {
    Write-Host "ERROR: FontRegister not found at:" -ForegroundColor Red
    Write-Host "  $FontRegisterPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Get current directory
$CurrentPath = Get-Location

# Find all font files
Write-Host "Scanning for fonts..." -ForegroundColor Cyan
$FontFiles = Get-ChildItem -Path $CurrentPath -Include "*.ttf", "*.otf" -File -Recurse

if ($FontFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "No font files found in current directory." -ForegroundColor Yellow
    Write-Host "Location: $CurrentPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 0
}

Write-Host "Found $($FontFiles.Count) font file(s)" -ForegroundColor Green
Write-Host ""

# Prepare font paths for installation
$FontPaths = $FontFiles | ForEach-Object { "`"$($_.FullName)`"" }

# Build command
$Arguments = @("install", "--machine") + $FontPaths

Write-Host "Installing fonts system-wide..." -ForegroundColor Cyan
Write-Host "─────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host ""

# Execute FontRegister
try {
    # Create temporary files for output
    $tempOutput = [System.IO.Path]::GetTempFileName()
    $tempError = [System.IO.Path]::GetTempFileName()
    
    $process = Start-Process -FilePath $FontRegisterPath `
                            -ArgumentList $Arguments `
                            -NoNewWindow `
                            -Wait `
                            -PassThru `
                            -RedirectStandardOutput $tempOutput `
                            -RedirectStandardError $tempError
    
    # Read output
    $output = Get-Content $tempOutput -Raw -ErrorAction SilentlyContinue
    $errorOut = Get-Content $tempError -Raw -ErrorAction SilentlyContinue
    
    # Display output if any
    if ($output) {
        Write-Host $output -ForegroundColor Gray
    }
    if ($errorOut) {
        Write-Host $errorOut -ForegroundColor Yellow
    }
    
    # Clean up temp files
    Remove-Item $tempOutput -Force -ErrorAction SilentlyContinue
    Remove-Item $tempError -Force -ErrorAction SilentlyContinue
    
    if ($process.ExitCode -eq 0) {
        Write-Host "✓ Successfully installed $($FontFiles.Count) font(s)" -ForegroundColor Green
        Write-Host ""
        
        # List installed fonts
        Write-Host "Installed fonts:" -ForegroundColor White
        foreach ($font in $FontFiles) {
            Write-Host "  • $($font.Name)" -ForegroundColor Gray
        }
    } else {
        Write-Host "✗ Installation completed with errors (Exit Code: $($process.ExitCode))" -ForegroundColor Yellow
    }
} catch {
    Write-Host "✗ Installation failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host ""
Write-Host "─────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "Installation complete. Fonts are now available system-wide." -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")