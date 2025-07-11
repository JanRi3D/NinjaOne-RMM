#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Deploys taskbar layout for Pressespiegel workstations

.DESCRIPTION
    This script deploys a predefined taskbar layout with Pressespiegel shortcuts.
    It copies shortcuts from network share and applies the layout configuration.

.NOTES
    Author: Jan Ried
    Version: 1.0
    Date: July 9, 2025
    Requires: Administrator privileges
#>

# Restart Explorer to apply registry changes
$explorerProcesses = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
if ($explorerProcesses) {
    $explorerProcesses | Stop-Process -Force
}

Start-Sleep -Seconds 2
Start-Process -FilePath "explorer.exe"
Start-Sleep -Seconds 3

$xmlContent = @'
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\01_Fehlermail_schreiben.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\02_Hilfe_Anleitungen_Dokumentation.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\03_Outlook.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\04_PS-Redaktionsmodul.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\05_Firefox_fuer_Pressespiegel.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\06_Firefox.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\07_Kundenordner.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\08_Chrome.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\09_Notepad++.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\10_PDF-XChange.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\11_Screenshot.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\12_Screenshot_Auswahl.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\13_Screenshot_Auswahl_bearbeiten.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\14_Zwischenablage_ansehen.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\15_PS-Konfigurator.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\16_Sumo.lnk" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Install\Pressespiegel\17_Texterkennung.lnk" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
'@
try {
    # Create installation directories
    $installDir = "C:\Install"
    $presseSpiegelDir = "$installDir\Pressespiegel"
    $iconsDir = "$presseSpiegelDir\Icons"
    
    if (-not (Test-Path $installDir)) {
        New-Item -Path $installDir -ItemType Directory -Force | Out-Null
    }
    
    if (-not (Test-Path $presseSpiegelDir)) {
        New-Item -Path $presseSpiegelDir -ItemType Directory -Force | Out-Null
    }
    
    if (-not (Test-Path $iconsDir)) {
        New-Item -Path $iconsDir -ItemType Directory -Force | Out-Null
    }
    
    # Copy shortcuts from network share
    $sourceDir = "\\fileserver01\PS_Install\buerorechner_einrichten\Symbolleiste_PS_Rechner"
    $linkFiles = @(
        "01_Fehlermail_schreiben.lnk",
        "02_Hilfe_Anleitungen_Dokumentation.lnk",
        "03_Outlook.lnk",
        "04_PS-Redaktionsmodul.lnk",
        "05_Firefox_fuer_Pressespiegel.lnk",
        "06_Firefox.lnk",
        "07_Kundenordner.lnk",
        "08_Chrome.lnk",
        "09_Notepad++.lnk",
        "10_PDF-XChange.lnk",
        "11_Screenshot.lnk",
        "12_Screenshot_Auswahl.lnk",
        "13_Screenshot_Auswahl_bearbeiten.lnk",
        "14_Zwischenablage_ansehen.lnk",
        "15_PS-Konfigurator.lnk",
        "16_Sumo.lnk",
        "17_Texterkennung.lnk"
    )
    
    # Icon mappings based on the numbered list
    $iconMappings = @{
        "01_Fehlermail_schreiben.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\essentials\fehlermail\icon\warning.ico"
        "02_Hilfe_Anleitungen_Dokumentation.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\Wikibooks-help-icon.ico"
        "03_Outlook.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\Outlook.ico"
        "04_PS-Redaktionsmodul.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\redaktionsmodul.ico"
        "05_Firefox_fuer_Pressespiegel.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\essentials\Firefox\icons\redaktion.ico"
        "11_Screenshot.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\screenshot.ico"
        "12_Screenshot_Auswahl.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\screenshot_rechteck.ico"
        "13_Screenshot_Auswahl_bearbeiten.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\bild_bearbeiten.ico"
        "14_Zwischenablage_ansehen.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\zwischenablage.ico"
        "15_PS-Konfigurator.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\konfigurator.ico"
        "16_Sumo.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\buerorechner_einrichten\icons\sumo.ico"
        "17_Texterkennung.lnk" = "\\fileserver01\Pressespiegel\PS_Technik\essentials\ocr\icon\ocr.ico"
    }
    
    # Copy icons from network share
    Write-Output "Copying icons from network share..."
    foreach ($linkFile in $linkFiles) {
        if ($iconMappings.ContainsKey($linkFile)) {
            $sourceIconPath = $iconMappings[$linkFile]
            $iconFileName = Split-Path $sourceIconPath -Leaf
            $destinationIconPath = "$iconsDir\$iconFileName"
            
            if (Test-Path $sourceIconPath) {
                try {
                    Copy-Item -Path $sourceIconPath -Destination $destinationIconPath -Force
                    Write-Output "Copied icon: $iconFileName"
                } catch {
                    Write-Warning "Failed to copy icon $iconFileName`: $($_.Exception.Message)"
                }
            } else {
                Write-Warning "Icon not found: $sourceIconPath"
            }
        }
    }
    
    if (-not (Test-Path $sourceDir)) {
        Write-Error "Network share not accessible: $sourceDir"
        exit 1
    }
    
    # Copy and fix shortcuts
    Write-Output "Copying and updating shortcuts..."
    foreach ($linkFile in $linkFiles) {
        $sourcePath = "$sourceDir\$linkFile"
        $destinationPath = "$presseSpiegelDir\$linkFile"
        
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destinationPath -Force
            
            # Fix batch file shortcuts for taskbar compatibility and update icons
            try {
                $shell = New-Object -ComObject WScript.Shell
                $shortcut = $shell.CreateShortcut($destinationPath)
                $targetPath = $shortcut.TargetPath
                $targetExt = [System.IO.Path]::GetExtension($targetPath).ToLower()
                
                $needsUpdate = $false
                
                # Fix batch file shortcuts
                if ($targetExt -eq ".bat" -or $targetExt -eq ".cmd") {
                    $shortcut.TargetPath = "cmd.exe"
                    $shortcut.Arguments = "/c `"$targetPath`""
                    $needsUpdate = $true
                }
                
                # Update icon if mapping exists
                if ($iconMappings.ContainsKey($linkFile)) {
                    $sourceIconPath = $iconMappings[$linkFile]
                    $iconFileName = Split-Path $sourceIconPath -Leaf
                    $localIconPath = "$iconsDir\$iconFileName"
                    
                    if (Test-Path $localIconPath) {
                        $shortcut.IconLocation = $localIconPath
                        $needsUpdate = $true
                        Write-Output "Updated icon for $linkFile"
                    }
                }
                
                if ($needsUpdate) {
                    $shortcut.Save()
                }
            } catch {
                Write-Warning "Failed to update shortcut $linkFile`: $($_.Exception.Message)"
            }
        } else {
            Write-Warning "Shortcut not found: $sourcePath"
        }
    }
    
    # Deploy taskbar layout
    $xmlPath = "$installDir\TaskbarLayout_Fixed.xml"
    $xmlContent | Out-File -FilePath $xmlPath -Encoding UTF8 -Force
    
    # Apply layout via registry
    $layoutRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
    if (-not (Test-Path $layoutRegPath)) {
        New-Item -Path $layoutRegPath -Force | Out-Null
    }
    
    Set-ItemProperty -Path $layoutRegPath -Name "StartLayoutFile" -Value $xmlPath -Type String
    Set-ItemProperty -Path $layoutRegPath -Name "LockedStartLayout" -Value 1 -Type DWord
    
    # Restart Explorer
    $explorerProcesses = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
    if ($explorerProcesses) {
        $explorerProcesses | Stop-Process -Force
    }
    
    Start-Sleep -Seconds 2
    Start-Process -FilePath "explorer.exe"
    Start-Sleep -Seconds 3
    exit 0
    
} catch {
    Write-Error "Failed to deploy taskbar layout: $($_.Exception.Message)"
    exit 1
}