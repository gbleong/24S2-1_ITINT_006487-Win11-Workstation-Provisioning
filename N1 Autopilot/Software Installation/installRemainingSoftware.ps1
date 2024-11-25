<# Initialize Global Variables #>

# Get the current folder path where the script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Get device manufacturer information
$deviceManufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer

<# Start Installation Processes #>

# Change working directory to script directory
Set-Location -Path $scriptDir

# > Install HP Support Assistant

# Check that device manufacturer is HP before installing HP software update application
if ($deviceManufacturer -like "HP*") {

    Start-Process -FilePath "sp148716.exe" -ArgumentList /s, /a, /s -Wait
}

# > Install GlobalProtect
Start-Process -FilePath "GlobalProtect64-6.0.3.msi" -ArgumentList /passive -Wait
#Start-Process msiexec.exe -ArgumentList "/f a GlobalProtect64-6.0.3.msi /passive" -Wait