<# Initialize Global Variables and Functions #>

# Get the current folder path where installers are located
$installersDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Installers"

# Get the current folder path where utility tools are located
$utilityToolsDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Utility Tools"

# Get the current folder path where credential files are located
$credentialFilesDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Credential Files"

# Get the folder path for temporary local app data files
$locAppDataTmpDir = $env:Temp

# Get device manufacturer information
$deviceManufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer

# Get device model information
$deviceModel = (Get-WmiObject Win32_ComputerSystem).Model

# Get device service tag from Win32_BIOS WMI class
$serviceTag = (Get-WmiObject Win32_BIOS).SerialNumber

# Input deployed Dell device models that uses Dell Encryption for storage drives
$ddsDeviceModels = @("Latitude 5300", "Latitude 5310", "Precision 7750")

# Function to check for internet connectivity
function checkInternetConnection {

    try {

        # Test connection using Google's DNS server
        $response = Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet
        return $response
    } 

    catch {

        # If an exception occurs, assume no connection
        return $false
    }
}



<# Start Software Installation Processes #>

# Change working directory to location of installers
Set-Location -Path $installersDir

# > Install Intel bluetooth driver

Start-Process -FilePath "BT-23.60.0-64UWD-Win10-Win11.exe" -ArgumentList /passive

# > Install Intel wi-fi driver

Start-Process -FilePath "WiFi-23.60.1-Driver64-Win10-Win11.exe" -ArgumentList /passive

# > Install Intel network connections driver

# Get list of folders in directory for temporary local app data files
$locAppDataTmpFolders = Get-ChildItem -Path $locAppDataTmpDir -Directory -ErrorAction SilentlyContinue

# Check for and remove leftover temporary WinZip Self-Extractor files
foreach ($folder in $locAppDataTmpFolders) {

    if ($folder.Name -match "^WZSE\d+\.TMP$") {

        Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Start-Process -FilePath "Wired_driver_29.3_x64.exe"

$targetWzseTmpFolder = $null

do {

    # Pause to allow WinZip Self-Extractor process to finish running
    Start-Sleep -Seconds 2

    # Get list of folders in directory for temporary local app data files
    $locAppDataTmpFolders = Get-ChildItem -Path $locAppDataTmpDir -Directory -ErrorAction SilentlyContinue

    # Check for and find folder matching the naming convention for WinZip Self-Extractor folders
    foreach ($folder in $locAppDataTmpFolders) {

        if ($folder.Name -match "^WZSE\d+\.TMP$") {

            # Store existing WinZip Self-Extractor folder name
            $targetWzseTmpFolder = $folder.Name

	        do {

                # Pause to allow WinZip Self-Extractor process to finish running
                Start-Sleep -Seconds 2

    		    # Construct the expected file path for target executable
                $setupbdExe = Join-Path $locAppDataTmpDir (Join-Path $targetWzseTmpFolder "APPS\SETUP\SETUPBD\Winx64\SetupBD.exe")

                # Run target file if it exists
                if (Test-Path -Path $setupbdExe) {

	                $process = Start-Process -FilePath $setupbdExe -ArgumentList "/switch /m" -PassThru
                    $process.WaitForExit()
                    break
                }

            } while (-not (Test-Path -Path $setupbdExe))

            break
        }
    }
    
} while (-not $targetWzseTmpFolder)

Stop-Process -Name "SetupBD" -Force -ErrorAction SilentlyContinue

# > Install Lenovo System Update

# Check that device manufacturer is Lenovo before installing Lenovo-specific software update application
if ($deviceManufacturer -like "LENOVO*") {

    Start-Process -FilePath "system_update_5.08.02.25.exe" -ArgumentList /SP-, /SILENT
}

# > Install HP Support Assistant

# Check that device manufacturer is HP before installing HP software update application
if ($deviceManufacturer -like "HP*") {

    Start-Process -FilePath "sp148716.exe" -ArgumentList /s, /a, /s
}

# > Install Google Chrome

if (-not (checkInternetConnection)) {

    Write-Host "Waiting for internet connection...`n"

    while (-not (checkInternetConnection)) {

        Start-Sleep -Seconds 4
    }

    Write-Host "Internet connection established. Continuing with installation...`n"
}

Start-Process -FilePath "ChromeSetup.exe" -ArgumentList "/silent /install"

# > Install Microsoft Office 365

Start-Process -FilePath "M365 Offline Installer 12-3-2024\setup.exe" -ArgumentList '/configure "M365 Offline Installer 12-3-2024\Configuration365.xml"'

# > Install Dell Command Update

# Check that device manufacturer is Dell before installing Dell-specific software update application
if ($deviceManufacturer -like "Dell*") {

    Start-Process -FilePath "Dell-Command-Update-Application_T45GH_WIN_5.3.0_A00.exe" -ArgumentList /passthrough, /v"/passive" -Wait
}

# > Install Comber

Start-Process -FilePath "Comber x64 1.0.1.0.exe" -ArgumentList /v"/passive" -Wait

# > Install Adobe Acrobat

Start-Process -FilePath "AcroRdrDCx642200320282_MUI.exe" -ArgumentList /sPB -Wait

# > Install Dell Encryption software

# Check that Dell device model is compatible before installing Dell Encryption software
if ($ddsDeviceModels -contains $deviceModel) {

    # Add Dell Encryption entitlement registry entry silently
    Start-Process regedit.exe -ArgumentList '/s "Dell Data Encryption\Dell Encryption Entitlement.reg"' -Wait

    # Input temporary directory path for storing Dell Encryption software installer files
    $tempDdsInstallersPath = Join-Path $locAppDataTmpDir "DDS Installers"

    # Clean up existing folders before creating a new temporary directory for installer files
    Remove-Item -Path $tempDdsInstallersPath -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $tempDdsInstallersPath -Force | Out-Null

    # Extract the Dell Encryption setup files into the temporary directory, and install all necessary software
    Start-Process -FilePath "Dell Data Encryption\DDSSetup.exe" -ArgumentList /s, /z"`"EXTRACT_INSTALLERS=$tempDdsInstallersPath`"" -Wait
    Start-Process -FilePath "$tempDdsInstallersPath\Prerequisites\Prereq_64bit_setup.exe" -ArgumentList /v"/passive" -Wait
    Start-Process -FilePath "$tempDdsInstallersPath\Encryption Management Agent\EMAgent_64bit_setup.exe" -ArgumentList '/v"/passive /norestart"' -Wait

    # Delete temporary folder for installer files
    Remove-Item -Path $tempDdsInstallersPath -Recurse -Force -ErrorAction SilentlyContinue
}



<# Start System Configuration Processes #>

# > Set Adobe Acrobat as default app for .pdf files

# Define path for temporary XML file
$tempXmlFile = "$locAppDataTmpDir\AppAssoc.xml"

# Clean up any existing folders before proceeding
if (Test-Path -Path $tempXmlFile) {

    Remove-Item -Path $tempXmlFile -Force -ErrorAction SilentlyContinue
}

# Export current default app associations into temporary XML file
dism.exe /online /Export-DefaultAppAssociations:$tempXmlFile | Out-Null

# Read exported XML file into an XML object for modification
$xmlContent = [xml] (Get-Content -Path $tempXmlFile -ErrorAction SilentlyContinue)

# Locate association node corresponding to files with the .pdf extension
$pdfAssociationNode = $xmlContent.SelectSingleNode("//Association[@Identifier='.pdf']")

# Update ProgId and ApplicationName variables in .pdf associate node to specify Adobe Acrobat as default app for .pdf files
$pdfAssociationNode.ProgId = "Acrobat.Document.DC"
$pdfAssociationNode.ApplicationName = "Adobe Acrobat"

# Save changes back to temporary XML file
$xmlContent.Save($tempXmlFile)

# Import updated default app associations from temporary XML file
dism.exe /online /Import-DefaultAppAssociations:$tempXmlFile | Out-Null

# Delete temporary XML file
Remove-Item -Path $tempXmlFile -Force -ErrorAction SilentlyContinue

# Change working directory to location of credential files
Set-Location -Path $credentialFilesDir

# > Create Windows local admin user accounts WIP

$n1WinLocalAdminCredFile = "N1 Windows Local Admin Credentials.csv"

if (Test-Path $n1WinLocalAdminCredFile) {

    $n1AdminAccounts = Import-Csv -Path $n1WinLocalAdminCredFile

    foreach ($account in $n1AdminAccounts) {

        $username = $account.Username
        $password = $account.Password

        # Create user account
        New-LocalUser -Name $username -password (ConvertTo-SecureString $password -AsPlainText -Force) -UserMayNotChangePassword -PasswordNeverExpires -AccountNeverExpires

        # Add the user to the Administrators group
        Add-LocalGroupMember -Group "Administrators" -Member $username
    }
} 

else {

    Write-Host "Windows local admin user accounts creation failed. Process will be skipped."
}



<# Start Device Profiling Processes #>

# Change working directory to location of utility tools
Set-Location -Path $utilityToolsDir

# > Display device service tag
Write-Host "Service Tag: $serviceTag`n"

# > Display device network adapter information

# Query WMI Win32_NetworkAdapter class for relevant information, store it as a string in a variable, and output it
$netAdapterInfo = Get-WmiObject Win32_NetworkAdapter | Select-Object NetConnectionID, Name, MACAddress | Out-String
Write-Host "Network Adapters:`n$netAdapterInfo"

# > Display storage drive serial number

# Open CrystalDiskInfo utility
Start-Process -FilePath "CrystalDiskInfo8_12_0\DiskInfo64.exe" -Wait

<#
# Pause to keep the console open
Write-Host "`nPress any key to exit..."
[System.Console]::ReadKey() | Out-Null
#>