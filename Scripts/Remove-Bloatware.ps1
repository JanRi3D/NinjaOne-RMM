#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Removes unwanted bloatware applications from Windows systems.

.DESCRIPTION
    This script removes pre-installed applications (bloatware) that are commonly unwanted
    on Windows systems. It targets both user-installed and system-provisioned packages
    to provide a cleaner Windows experience.

.EXAMPLE
    .\Remove-Bloatware.ps1
    
    Removes all predefined bloatware applications from the system.

.NOTES
    Author: Jan Ried
    Version: 1.0
    Date: July 8, 2025
    Requires: PowerShell 5.1+ and Administrator privileges
    
    This script will remove applications for all users and prevent reinstallation
    through provisioned packages. Use with caution in enterprise environments.
#>

# Function to remove AppX packages
function Remove-AppXPackage {
    param([string]$PackageName)
    
    try {
        # Remove for current user
        Get-AppxPackage -Name $PackageName -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
        
        # Remove for all users
        Get-AppxPackage -AllUsers -Name $PackageName -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        
        # Remove provisioned package to prevent reinstallation
        Get-AppxProvisionedPackage -Online | Where-Object DisplayName -eq $PackageName | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
        
        Write-Host "Successfully removed: $PackageName" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to remove $PackageName : $($_.Exception.Message)"
    }
}

# List of bloatware applications to remove
$bloatwareApps = @(
    "Microsoft.3DBuilder",
    "Microsoft.BingFinance",
    "Microsoft.BingNews",
    "Microsoft.BingSports",
    "Microsoft.BingWeather",
    "Microsoft.Getstarted",
    "Microsoft.MicrosoftOfficeHub",
    "Microsoft.MicrosoftSolitaireCollection",
    "Microsoft.People",
    "Microsoft.SkypeApp",
    "Microsoft.WindowsAlarms",
    "Microsoft.windowscommunicationsapps",
    "Microsoft.WindowsFeedbackHub",
    "Microsoft.WindowsMaps",
    "Microsoft.WindowsSoundRecorder",
    "Microsoft.Xbox.TCUI",
    "Microsoft.XboxApp",
    "Microsoft.XboxGameOverlay",
    "Microsoft.XboxIdentityProvider",
    "Microsoft.XboxSpeechToTextOverlay",
    "Microsoft.ZuneMusic",
    "Microsoft.ZuneVideo"
)

# Remove each bloatware application
foreach ($app in $bloatwareApps) {
    Remove-AppXPackage -PackageName $app
}

Write-Host "Bloatware removal completed!" -ForegroundColor Cyan