#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs applications using Windows Package Manager (WinGet).

.DESCRIPTION
    This script automatically installs a predefined list of applications using WinGet.
    It locates the WinGet executable from the Microsoft.DesktopAppInstaller package
    and installs each application with automatic acceptance of license agreements.

.EXAMPLE
    .\WinGet.ps1
    
    Installs all applications defined in the $apps array.

.NOTES
    Author: NinjaOne RMM Script
    Version: 1.0
    Date: July 8, 2025
    Requires: PowerShell 5.1+, Administrator privileges, and WinGet (Microsoft.DesktopAppInstaller)
    
    Modify the $apps array to include the desired application IDs.
    Use 'winget search <appname>' to find the correct application IDs.
#>

$ErrorActionPreference = 'Stop'

function Get-WinGetExecutable {
    $winget = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq 'Microsoft.DesktopAppInstaller' }
    if ($null -ne $winget) {
        return Join-Path -Path $winget.InstallLocation -ChildPath 'winget.exe'
    } else {
        Write-Error 'WinGet nicht gefunden. Microsoft.DesktopAppInstaller fehlt oder ist nicht installiert.'
        exit 1
    }
}

$wingetExe = Get-WinGetExecutable

$apps = @(
    'Your APP ID 1'
    'Your APP ID 2'
)

foreach ($app in $apps) {
    Write-Host "Installing $app"
    & $wingetExe install --id $app -e --accept-source-agreements --accept-package-agreements
}