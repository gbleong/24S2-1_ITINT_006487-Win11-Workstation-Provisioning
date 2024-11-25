<# Initialize Global Variables #>

# Get the current folder path where the script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

<# Start Profiling Processes #>

# Change working directory to script directory
Set-Location -Path $scriptDir

# Query the Win32_BIOS WMI class to get the SerialNumber (service tag)
$ServiceTag = (Get-WmiObject -Class Win32_BIOS).SerialNumber

# Output the Service Tag (optional)
Write-Host "The Service Tag of this computer is: $ServiceTag"

# Get the LAN Ethernet MAC address and store it in a variable

$EthernetMACAddress = Get-CimInstance -ClassName Win32_NetworkAdapter |
    Where-Object {
        $_.NetConnectionID -eq "Ethernet" -and $_.MACAddress -ne $null
    } |
    Select-Object -ExpandProperty MACAddress -First 1

# Check if a MAC address was found and display a message
if ($EthernetMACAddress) {
    Write-Host "The LAN Ethernet MAC address is: $EthernetMACAddress"
} else {
    Write-Host "No LAN Ethernet adapter found with a MAC address."
}

Start-Process -FilePath "DiskInfo64.exe" -Wait

<#
# Path to the file containing the drive information
$filePath = ".\cdiSample.txt"

# Initialize variables to store serial numbers of internal drives
$storagedrives = @()

# Read the file line by line
$fileContent = Get-Content -Path $filePath

# Variables to track drive type and serial number
$currentDriveType = ""
$currentSerial = ""

foreach ($line in $fileContent) {
    # Check if the line contains information about a new drive
    if ($line -match "^\s*\(\d+\)\s+(.*?)\s+:") {
        # When starting a new drive, reset current drive type and serial
        $currentDriveType = ""
        $currentSerial = ""
    }

    # Check for drive serial number
    if ($line -match "^\s*Serial Number\s+:\s+(.*)") {
        $currentSerial = $matches[1]
    }

    # Check for USB or internal indicator
    if ($line -match "^\s*Interface\s+:\s+(.*)") {
        $interface = $matches[1]
        if ($interface -match "UASP|USB") {
            $currentDriveType = "External"
        } else {
            $currentDriveType = "Internal"
        }
    }

    # Process internal drive information
    if ($currentDriveType -eq "Internal" -and $currentSerial -ne "") {
        $storagedrives += $currentSerial
        # Clear current drive type to prevent duplicates
        $currentDriveType = ""
    }
}

# Store the serials in variables dynamically
for ($i = 0; $i -lt $storagedrives.Count; $i++) {
    Set-Variable -Name "storagedrive$($i + 1)" -Value $storagedrives[$i]
}

# Output the variables for verification
Write-Host "Internal Storage Drives and their Serial Numbers:"
for ($i = 0; $i -lt $storagedrives.Count; $i++) {
    Write-Host "storagedrive$($i + 1): $($storagedrives[$i])"
}
#>