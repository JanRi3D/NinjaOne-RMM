Stop-Service wuauserv -Force
Stop-Service bits -Force
Stop-Service cryptsvc -Force
Stop-Service trustedinstaller -Force

Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations" -Force -ErrorAction SilentlyContinue

Remove-Item "C:\Windows\SoftwareDistribution\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "C:\Windows\System32\catroot2\*" -Recurse -Force -ErrorAction SilentlyContinue

Start-Service trustedinstaller
Start-Service cryptsvc
Start-Service bits
Start-Service wuauserv


Import-Module PSWindowsUpdate
Get-WindowsUpdate -AcceptAll -Install -AutoReboot
