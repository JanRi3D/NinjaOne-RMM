#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Comprehensive Windows system update script for all major update sources.

.DESCRIPTION
    This script performs a complete system update by managing updates from multiple sources:
    - Windows Updates (using PSWindowsUpdate module)
    - Application updates via Windows Package Manager (WinGet)
    - Application updates via Chocolatey package manager
    
    The script includes comprehensive logging, error handling, and distinct exit codes
    for troubleshooting. Designed for deployment via NinjaOne RMM systems.

.PARAMETER LogPath
    The path where the update log file will be created. 
    Default is "$env:TEMP\UpdateScript.log".

.PARAMETER SkipWindowsUpdates
    Skip Windows Update installation and only update applications.

.PARAMETER SkipWinGet
    Skip WinGet application updates.

.PARAMETER SkipChocolatey
    Skip Chocolatey application updates.

.EXAMPLE
    .\Windows-Update.ps1
    
    Runs all update processes with default logging location.

.EXAMPLE
    .\Windows-Update.ps1 -LogPath "C:\Logs\Updates.log"
    
    Runs all updates with custom log file location.

.EXAMPLE
    .\Windows-Update.ps1 -SkipWindowsUpdates
    
    Only updates applications via WinGet and Chocolatey, skips Windows Updates.

.NOTES
    Author: NinjaOne RMM Script
    Version: 1.0
    Date: July 9, 2025
    Requires: PowerShell 5.1+ and Administrator privileges
    
    Exit Codes:
    0 = Success
    1 = Windows Update failure
    2 = WinGet update failure
    3 = Chocolatey update failure
    
    The script automatically installs the PSWindowsUpdate module if not present.
    For enterprise environments, consider pre-installing required modules.
#>

param (
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "$env:TEMP\UpdateScript.log",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipWindowsUpdates,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipWinGet,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipChocolatey
)

# Function to write timestamped log entries
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    
    # Write to both log file and console
    $logEntry | Out-File -FilePath $LogPath -Append -Encoding utf8
    
    switch ($Level) {
        "INFO" { Write-Host $logEntry -ForegroundColor Green }
        "WARN" { Write-Warning $logEntry }
        "ERROR" { Write-Error $logEntry }
    }
}

# Function to check if WinGet is available
function Test-WinGetAvailable {
    try {
        $winget = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq 'Microsoft.DesktopAppInstaller' }
        
        if ($null -ne $winget) {
            $wingetPath = Join-Path -Path $winget.InstallLocation -ChildPath 'winget.exe'
            return Test-Path $wingetPath
        }
        return $false
    }
    catch {
        return $false
    }
}

# Function to install Windows Updates
function Install-WindowsUpdates {
    try {
        Write-Log "Checking for PSWindowsUpdate module..."
        
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Log "PSWindowsUpdate module not found. Installing..."
            Install-Module -Name PSWindowsUpdate -Force -Confirm:$false -ErrorAction Stop
            Write-Log "PSWindowsUpdate module installed successfully."
        }
        
        Write-Log "Importing PSWindowsUpdate module..."
        Import-Module PSWindowsUpdate -ErrorAction Stop
        
        Write-Log "Starting Windows Updates scan and installation..."
        $updates = Get-WindowsUpdate -AcceptAll -Install -AutoReboot -IgnoreReboot -ErrorAction Stop
        
        if ($updates) {
            $updates | ForEach-Object { 
                Write-Log "Installed: $($_.Title) - $($_.Result)"
            }
        } else {
            Write-Log "No Windows Updates available."
        }
        
        Write-Log "Windows Updates process completed successfully."
        return $true
    }
    catch {
        Write-Log "Windows Update process failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to update applications via WinGet
function Update-WinGetApplications {
    try {
        Write-Log "Checking WinGet availability..."
        
        if (-not (Test-WinGetAvailable)) {
            Write-Log "WinGet not found or not available." "WARN"
            return $true
        }
        
        Write-Log "Starting WinGet application updates..."
        $wingetOutput = winget upgrade --all --silent --accept-source-agreements --accept-package-agreements 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            $wingetOutput | ForEach-Object { Write-Log "WinGet: $_" }
            Write-Log "WinGet application updates completed successfully."
            return $true
        } else {
            Write-Log "WinGet update process returned exit code: $LASTEXITCODE" "WARN"
            $wingetOutput | ForEach-Object { Write-Log "WinGet Error: $_" "WARN" }
            return $true # Don't fail the entire script for WinGet issues
        }
    }
    catch {
        Write-Log "WinGet update process failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to update applications via Chocolatey
function Update-ChocolateyApplications {
    try {
        Write-Log "Checking Chocolatey availability..."
        
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
            Write-Log "Chocolatey not found." "WARN"
            return $true
        }
        
        Write-Log "Starting Chocolatey application updates..."
        $chocoOutput = choco upgrade all -y 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            $chocoOutput | ForEach-Object { Write-Log "Chocolatey: $_" }
            Write-Log "Chocolatey application updates completed successfully."
            return $true
        } else {
            Write-Log "Chocolatey update process returned exit code: $LASTEXITCODE" "WARN"
            $chocoOutput | ForEach-Object { Write-Log "Chocolatey Error: $_" "WARN" }
            return $true # Don't fail the entire script for Chocolatey issues
        }
    }
    catch {
        Write-Log "Chocolatey update process failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Main execution
try {
    Write-Log "Windows Update Script started."
    Write-Log "Log file location: $LogPath"
    
    $overallSuccess = $true
    
    # Windows Updates
    if (-not $SkipWindowsUpdates) {
        Write-Log "Processing Windows Updates..."
        if (-not (Install-WindowsUpdates)) {
            Write-Log "Windows Updates failed." "ERROR"
            $overallSuccess = $false
            exit 1
        }
    } else {
        Write-Log "Skipping Windows Updates as requested."
    }
    
    # WinGet Updates
    if (-not $SkipWinGet) {
        Write-Log "Processing WinGet application updates..."
        if (-not (Update-WinGetApplications)) {
            Write-Log "WinGet updates failed." "ERROR"
            $overallSuccess = $false
            exit 2
        }
    } else {
        Write-Log "Skipping WinGet updates as requested."
    }
    
    # Chocolatey Updates
    if (-not $SkipChocolatey) {
        Write-Log "Processing Chocolatey application updates..."
        if (-not (Update-ChocolateyApplications)) {
            Write-Log "Chocolatey updates failed." "ERROR"
            $overallSuccess = $false
            exit 3
        }
    } else {
        Write-Log "Skipping Chocolatey updates as requested."
    }
    
    if ($overallSuccess) {
        Write-Log "All update processes completed successfully."
        exit 0
    } else {
        Write-Log "Some update processes encountered issues. Check the log for details." "WARN"
        exit 1
    }
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" "ERROR"
    exit 1
}
