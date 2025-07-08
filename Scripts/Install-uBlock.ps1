#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs uBlock Origin for Firefox and uBlock Lite for Chrome and Edge browsers.

.DESCRIPTION
    This script automatically installs ad-blocking extensions for the three major browsers:
    - Firefox: uBlock Origin
    - Chrome: uBlock Lite  
    - Edge: uBlock Lite
    
    The script handles registry modifications and policy configurations to enable the extensions.

.NOTES
    Author: NinjaOne RMM Script
    Version: 1.0
    Date: July 8, 2025
    Requires: PowerShell 5.1+ and Administrator privileges
#>

# Function to check if a browser is installed
function Test-BrowserInstalled {
    param([string]$BrowserName)
    
    switch ($BrowserName) {
        "Firefox" {
            return (Test-Path "C:\Program Files\Mozilla Firefox\firefox.exe") -or 
                   (Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe")
        }
        "Chrome" {
            return (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") -or
                   (Test-Path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
        }
        "Edge" {
            return (Test-Path "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") -or
                   (Get-AppxPackage -Name "Microsoft.MicrosoftEdge*" -ErrorAction SilentlyContinue)
        }
    }
    return $false
}

# Function to install uBlock Origin for Firefox
function Install-FirefoxUBlock {
    try {
        # Firefox extension ID and download URL
        $extensionId = "uBlock0@raymondhill.net"
        $downloadUrl = "https://addons.mozilla.org/firefox/downloads/latest/ublock-origin/addon-607454-latest.xpi"
        
        # Create Firefox policies directory
        $policiesPath = "C:\Program Files\Mozilla Firefox\distribution"
        if (!(Test-Path $policiesPath)) {
            New-Item -Path $policiesPath -ItemType Directory -Force | Out-Null
        }
        
        # Create policies.json for Firefox
        $policiesJson = @{
            policies = @{
                Extensions = @{
                    Install = @($downloadUrl)
                }
                ExtensionSettings = @{
                    $extensionId = @{
                        installation_mode = "force_installed"
                        install_url = $downloadUrl
                    }
                }
            }
        } | ConvertTo-Json -Depth 10
        
        Set-Content -Path "$policiesPath\policies.json" -Value $policiesJson -Encoding UTF8
        return $true
    }
    catch {
        Write-Error "Failed to configure uBlock Origin for Firefox: $($_.Exception.Message)"
        return $false
    }
}

# Function to install uBlock Lite for Chrome
function Install-ChromeUBlockLite {
    try {
        # Chrome extension ID for uBlock Lite
        $extensionId = "ddkjiahejlhfcafbddmgiahcphecmpfh"
        
        # Registry path for Chrome policies
        $regPath = "HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist"
        
        # Create registry path if it doesn't exist
        if (!(Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        # Add extension to force install list
        $extensionUrl = "$extensionId;https://clients2.google.com/service/update2/crx"
        
        # Find next available registry value
        $existingValues = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        $nextIndex = 1
        if ($existingValues) {
            $numbers = ($existingValues.PSObject.Properties | Where-Object { $_.Name -match '^\d+$' }).Name | ForEach-Object { [int]$_ }
            if ($numbers) {
                $nextIndex = ($numbers | Measure-Object -Maximum).Maximum + 1
            }
        }
        
        New-ItemProperty -Path $regPath -Name $nextIndex -Value $extensionUrl -PropertyType String -Force | Out-Null
        return $true
    }
    catch {
        Write-Error "Failed to configure uBlock Lite for Chrome: $($_.Exception.Message)"
        return $false
    }
}

# Function to install uBlock Lite for Edge
function Install-EdgeUBlockLite {
    try {
        # Edge extension ID for uBlock Lite
        $extensionId = "cjpalhdlnbpafiamejdnhcphjbkeiagm"
        
        # Registry path for Edge policies
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist"
        
        # Create registry path if it doesn't exist
        if (!(Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        # Add extension to force install list
        $extensionUrl = "$extensionId;https://edge.microsoft.com/extensionwebstorebase/v1/crx"
        
        # Find next available registry value
        $existingValues = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        $nextIndex = 1
        if ($existingValues) {
            $numbers = ($existingValues.PSObject.Properties | Where-Object { $_.Name -match '^\d+$' }).Name | ForEach-Object { [int]$_ }
            if ($numbers) {
                $nextIndex = ($numbers | Measure-Object -Maximum).Maximum + 1
            }
        }
        
        New-ItemProperty -Path $regPath -Name $nextIndex -Value $extensionUrl -PropertyType String -Force | Out-Null
        return $true
    }
    catch {
        Write-Error "Failed to configure uBlock Lite for Edge: $($_.Exception.Message)"
        return $false
    }
}

# Main execution
$successCount = 0
$totalBrowsers = 0

# Check and install for Firefox
if (Test-BrowserInstalled "Firefox") {
    $totalBrowsers++
    if (Install-FirefoxUBlock) {
        $successCount++
    }
}

# Check and install for Chrome
if (Test-BrowserInstalled "Chrome") {
    $totalBrowsers++
    if (Install-ChromeUBlockLite) {
        $successCount++
    }
}

# Check and install for Edge
if (Test-BrowserInstalled "Edge") {
    $totalBrowsers++
    if (Install-EdgeUBlockLite) {
        $successCount++
    }
}

# Return results
if ($totalBrowsers -eq 0) {
    Write-Warning "No supported browsers found on this system."
    exit 1
} elseif ($successCount -ne $totalBrowsers) {
    Write-Warning "Some extensions failed to configure."
    exit 1
} else {
    Write-Output "All extensions configured successfully. Restart browsers to apply changes."
    exit 0
}