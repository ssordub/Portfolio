@echo off
Rem Make powershell read this file, skip a number of lines, and execute it.
Rem This works around .ps1 bad file association as non executables.
PowerShell -Command "Get-Content '%~dpnx0' | Select-Object -Skip 5 | Out-String | Invoke-Expression"
goto :eof

function Get-HardwareDevices {
    Get-WmiObject -Class Win32_PnPEntity | Format-Table Name, Manufacturer, DeviceID -AutoSize
    Write-Host "Press any key to continue ..."
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Get-HardwareDevices