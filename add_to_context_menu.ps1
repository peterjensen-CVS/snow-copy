#### This script will add a context menu item through HKEY Current User
#### Doing so will allow it to not need admin rights
#### The script Path is presumed

### Written by Peter Jensen


$menuName = "Open with ServiceNow"
$scriptPath = "C:\Snow-Copy\snow-copy.ps1"

# Set the registry path
$regPath = "HKCU:\Software\Classes\Directory\Background\shell\$menuName"
$commandPath = "$regPath\command"

# Create registry entries
New-Item -Path $regPath -Force
Set-ItemProperty -Path $regPath -Name "Icon" -Value "msedge.exe"

New-Item -Path $commandPath -Force
Set-ItemProperty -Path $commandPath -Name "(default)" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""

Write-Output "Context menu item '$menuName' added."
