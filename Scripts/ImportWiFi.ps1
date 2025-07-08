#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Imports and connects to a WiFi network profile on Windows systems.

.DESCRIPTION
    This script creates a WiFi profile XML file from a template and imports it into Windows.
    It supports WPA3-SAE authentication with AES encryption and automatically connects to the network.
    The script creates the necessary directory structure and cleans up temporary files after execution.

.PARAMETER WiFiName
    The name (SSID) of the WiFi network to configure. Default is "NAME".

.PARAMETER WiFiPassword
    The password for the WiFi network. Default is "PASSWORD".

.PARAMETER LocalPath
    The local path where the temporary WiFi profile XML will be created. 
    Default is "C:\Install\Wi-Fi\WiFi.xml".

.EXAMPLE
    .\ImportWiFi.ps1 -WiFiName "MyNetwork" -WiFiPassword "MyPassword123"
    
.EXAMPLE
    .\ImportWiFi.ps1 -WiFiName "CompanyWiFi" -WiFiPassword "SecurePass" -LocalPath "C:\Temp\profile.xml"

.NOTES
    Author: NinjaOne RMM Script
    Version: 1.0
    Date: July 8, 2025
    Requires: PowerShell 5.1+ and Administrator privileges
    
    The script uses WPA3-SAE with transition mode for backward compatibility.
    Temporary XML files are automatically cleaned up after import.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$WiFiName = "NAME",  # Change this to your desired WiFi name
   
    [Parameter(Mandatory=$false)]
    [string]$WiFiPassword = "PASSWORD",  # Change this to your desired password
   
    [Parameter(Mandatory=$false)]
    [string]$LocalPath = "C:\Install\Wi-Fi\WiFi.xml"
)

$wifiProfileTemplate = @'
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{WIFI_NAME}</name>
    <SSIDConfig>
        <SSID>
            <hex>{WIFI_HEX}</hex>
            <name>{WIFI_NAME}</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA3SAE</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
                <transitionMode xmlns="http://www.microsoft.com/networking/WLAN/profile/v4">true</transitionMode>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>{WIFI_PASSWORD}</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
'@

function Convert-StringToHex {
    param([string]$InputString)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($InputString)
    return ($bytes | ForEach-Object { $_.ToString("X2") }) -join ''
}

try {
    # Create directory if it doesn't exist
    $dir = Split-Path $LocalPath -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    
    # Convert WiFi name to hex and create profile
    $wifiHex = Convert-StringToHex -InputString $WiFiName
    $wifiProfile = $wifiProfileTemplate -replace '{WIFI_NAME}', $WiFiName
    $wifiProfile = $wifiProfile -replace '{WIFI_HEX}', $wifiHex
    $wifiProfile = $wifiProfile -replace '{WIFI_PASSWORD}', $WiFiPassword
    
    # Save and import profile
    $wifiProfile | Set-Content -Path $LocalPath -Encoding UTF8
    $result = netsh wlan add profile filename="$LocalPath" user=all
    
    if ($LASTEXITCODE -eq 0) {
        netsh wlan connect name="$WiFiName"
        Write-Host "WiFi profile '$WiFiName' imported and connected successfully!" -ForegroundColor Green
    } else {
        Write-Error "Failed to import WiFi profile. Error: $result"
    }
    
    # Clean up
    if (Test-Path $LocalPath) {
        Remove-Item $LocalPath -Force
    }
    
} catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}