<# Initialize Global Variables #>

# Get the current folder path where the script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Get device manufacturer information
$deviceManufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer

# Input VPN software name
$vpnSoftwareName = "GlobalProtect"

<# Start Installation Processes #>

# Change working directory to script directory
Set-Location -Path $scriptDir

# > Install HP Support Assistant

# Check that device manufacturer is HP before installing HP software update application
if ($deviceManufacturer -like "HP*") {

    Start-Process -FilePath "sp148716.exe" -ArgumentList /s, /a, /s
}

# > Install or repair GlobalProtect

# Query WMI to check if GlobalProtect is already installed, and install or repair it accordingly
$installCheck = Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -like "$vpnSofwareName" }

if ($installCheck) {

    Start-Process -FilePath msiexec.exe -ArgumentList "/fm GlobalProtect64-6.0.3.msi /passive" -Wait
} 

else {

    Start-Process -FilePath GlobalProtect64-6.0.3.msi -ArgumentList /passive -Wait
}