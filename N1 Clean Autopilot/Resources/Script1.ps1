<# Initialize Global Variables and Functions #>

# Get the current folder path where installers are located
$installersDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Installers"

# Get the current folder path where utility tools are located
$utilityToolsDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Utility Tools"

# Get the folder path for temporary local app data files
$locAppDataTmpDir = $env:Temp

# Get device manufacturer information
$deviceManufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer

# Get device service tag from Win32_BIOS WMI class
$serviceTag = (Get-WmiObject -Class Win32_BIOS).SerialNumber



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

        Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
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

Stop-Process -Name "SetupBD" -Force

# > Install Lenovo System Update

# Check that device manufacturer is Lenovo before installing Lenovo-specific software update application
if ($deviceManufacturer -like "LENOVO*") {

    Start-Process -FilePath "system_update_5.08.02.25.exe" -ArgumentList /SP-, /SILENT
}

# > Install Dell Command Update

# Check that device manufacturer is Dell before installing Dell-specific software update application
if ($deviceManufacturer -like "Dell*") {

    Start-Process -FilePath "Dell-Command-Update-Application_T45GH_WIN_5.3.0_A00.exe" -ArgumentList /passthrough, /v"/passive" -Wait
}

# > Install Comber

Start-Process -FilePath "Comber x64 1.0.1.0.exe" -ArgumentList /v"/passive" -Wait



<# Start System Configuration Processes #>

# > Set Adobe Acrobat as default app for .pdf files

# Define path for temporary XML file
$tempXmlFile = "$locAppDataTmpDir\AppAssoc.xml"

# Clean up any existing folders before proceeding
Remove-Item -Path $tempXmlFile -Force -ErrorAction SilentlyContinue

# Export current default app associations into temporary XML file
dism.exe /online /Export-DefaultAppAssociations:$tempXmlFile | Out-Null

# Read exported XML file into an XML object for modification
$xmlContent = [xml] (Get-Content -Path $tempXmlFile)

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



<# Start Device Profiling Processes #>

# Change working directory to location of utility tools
Set-Location -Path $utilityToolsDir

# > Display device service tag
Write-Host "Service Tag: $serviceTag`n"

# > Display device network adapter information

# Use the Get-CimInstance cmdlet to query network adapters for relevant information, store it as a string in a variable, and output it
$netAdapterInfo = Get-CimInstance -ClassName Win32_NetworkAdapter | Select-Object NetConnectionID, Name, MACAddress | Out-String
Write-Host "Network Adapters:`n$netAdapterInfo"

# > Display storage drive serial number

# Open CrystalDiskInfo utility
Start-Process -FilePath ".\CrystalDiskInfo8_12_0\DiskInfo64.exe" -Wait

<#
# Pause to keep the console open
Write-Host "`nPress any key to exit..."
[System.Console]::ReadKey() | Out-Null
#>