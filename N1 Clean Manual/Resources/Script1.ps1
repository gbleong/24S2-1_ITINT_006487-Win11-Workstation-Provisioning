<# Initialize Global Variables and Functions #>

# Get the current folder path where installers are located
$installersDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Installers"

# Get the current folder path where other files are located
$othersDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Others"

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
        $response = Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet -ErrorAction SilentlyContinue
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

    Start-Process -FilePath "system_update_5.08.02.25.exe" -ArgumentList "/SILENT /NORESTART"
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

    Start-Sleep -Seconds 2
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

Start-Process -FilePath "Comber x64 1.1.0.0.exe" -ArgumentList /v"/passive" -Wait

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
    Start-Process -FilePath "$tempDdsInstallersPath\Prerequisites\Prereq_64bit_setup.exe" -ArgumentList '/s /v"/passive"' -Wait
    Start-Process -FilePath "$tempDdsInstallersPath\Encryption Management Agent\EMAgent_64bit_setup.exe" -ArgumentList '/s /v"/passive /norestart"' -Wait

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
dism.exe /online /Export-DefaultAppAssociations:$tempXmlFile /quiet >$null 2>&1

# Read exported XML file into an XML object for modification
$xmlContent = [xml] (Get-Content -Path $tempXmlFile -ErrorAction SilentlyContinue)

# Locate association node corresponding to files with the .pdf extension
$pdfAssociationNode = $xmlContent.SelectSingleNode("//Association[@Identifier='.pdf']")

# Update ProgId and ApplicationName variables in .pdf associate node to specify Adobe Acrobat as default app for .pdf files
$pdfAssociationNode.ProgId = "Acrobat.Document.DC"
$pdfAssociationNode.ApplicationName = "Adobe Acrobat"

# Save changes back to temporary XML file
$xmlContent.OuterXml | Set-Content -Path $tempXmlFile -Force

# Import updated default app associations from temporary XML file
dism.exe /online /Import-DefaultAppAssociations:$tempXmlFile /quiet >$null 2>&1

# Delete temporary XML file
Remove-Item -Path $tempXmlFile -Force -ErrorAction SilentlyContinue

# > Turn off or disable privacy & security settings

# Define array of registry entries related to privacy and security settings
$privSecRegEntries = @(

    # Disables choose privacy settings experience for all users
    @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE"; Name = "DisablePrivacyExperience"; Type = "DWord"; Value = 1},

    # Turns off location services for all users
    @{Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"; Name = "Value"; Type = "String"; Value = "Deny"},

    # Turns off Find My Device services for all users
    @{Path = "HKLM:\SOFTWARE\Microsoft\MdmCommon\SettingValues"; Name = "LocationSyncEnabled"; Type = "DWord"; Value = 0},

    # Disables diagnostic data collection for all users
    @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"; Name = "AllowTelemetry"; Type = "DWord"; Value = 0},

    # Disables inking and typing improvement services for all users
    @{Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\TextInput"; Name = "AllowLinguisticDataCollection"; Type = "DWord"; Value = 0},

    # Disables advertising ID usage for all users
    @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo"; Name = "DisabledByGroupPolicy"; Type = "DWord"; Value = 1}
)

# Iterate over each registry entry, and create or modify it
foreach ($regEntry in $privSecRegEntries) {

    # Creates registry key if it does not already exist
    New-Item -Path $regEntry.Path -Force | Out-Null

    # Add the specified registry entry
    Set-ItemProperty -Path $regEntry.Path -Name $regEntry.Name -Value $regEntry.Value -Type $regEntry.Type
}

# Define path to default user NTUSER.DAT file
$defaultUserHivePath = "C:\Users\Default\NTUSER.DAT"

# Define temporrary key name for mounting registry hive
$tempHiveKeyName = "TempDefaultUserHive"

# # Define array of registry entries related to privacy and security settings
$privSecRegEntries = @(

    # Turns off improve inking and typing services in default user configuration profile
    @{Path = "Software\Microsoft\Input\TIPC"; Name = "Enabled"; Type = "REG_DWORD"; Value = 0},

    # Turns off advertising ID usage in default user configuration profile
    @{Path = "Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo"; Name = "Enabled"; Type = "REG_DWORD"; Value = 0},

    # Turns off tailored experiences services in default user configuration profile
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Privacy"; Name = "TailoredExperiencesWithDiagnosticDataEnabled"; Type = "REG_DWORD"; Value = 0},

    # Disables tailored experiences services in default user configuration profile
    @{Path = "Software\Policies\Microsoft\Windows\CloudContent"; Name = "DisableTailoredExperiencesWithDiagnosticData"; Type = "REG_DWORD"; Value = 1},

    # Turns off website access to language list in default user configuration profile
    @{Path = "Control Panel\International\User Profile"; Name = "HttpAcceptLanguageOptOut"; Type = "REG_DWORD"; Value = 1},

    # Turns off app launch tracking in default user configuration profile
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name = "Start_TrackProgs"; Type = "REG_DWORD"; Value = 0},

    # Disables app launch tracking in default user configuration profile
    @{Path = "Software\Policies\Microsoft\Windows\EdgeUI"; Name = "DisableMFUTracking"; Type = "REG_DWORD"; Value = 1},

    # Turns off suggested content in settings in default user configuration profile
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-338393Enabled"; Type = "REG_DWORD"; Value = 0},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-353694Enabled"; Type = "REG_DWORD"; Value = 0},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-353696Enabled"; Type = "REG_DWORD"; Value = 0},

    # Turn off notifications in settings app in default user configuration profile
    @{Path = "Software\Microsoft\Windows\CurrentVersion\SystemSettings\AccountNotifications"; Name = "EnableAccountNotifications"; Type = "REG_DWORD"; Value = 0}
)

# Load registry hive into temporary registry key
reg.exe load HKU\$tempHiveKeyName $defaultUserHivePath >$null 2>&1

# Create or modify the registry entries
foreach ($entry in $privSecRegEntries) {

    # Creates registry key if it does not already exist
    New-Item -Path "HKU\$tempHiveKeyName\$($entry.Path)" -Force | Out-Null

    # Add the specified registry entry
    reg.exe add `"HKU\$tempHiveKeyName\$($entry.Path)`" /v $($entry.Name) /t $($entry.Type) /d $($entry.Value) /f >$null 2>&1
}

# Unload registry hive from temporary registry key
reg.exe unload HKU\$tempHiveKeyName >$null 2>&1

# Change working directory to location of other files
Set-Location -Path $othersDir

# > Import Menlo security root certificate

# Open local machine certificate store for Trusted Root Certification Authorities
$certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")

# Create certificate object
$certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(Join-Path -Path (Get-Location) -ChildPath "MenloSecurityRootCA.crt")

# Import certificate to store
$certStore.Open('ReadWrite')
$certStore.Add($certificate)
$certStore.Close()

# Clean up and release resources
$certificate.Dispose()

# Change working directory to location of credential files
Set-Location -Path $credentialFilesDir

# > Create Windows local admin user accounts

# Import contents of the CSV file containing Windows local admin credentials as an array of objects where each row represents an account
$n1AdminAccounts = Import-Csv -Path "N1 Windows Local Admin Credentials.csv"

 # Iterate through each account object in the imported CSV
foreach ($account in $n1AdminAccounts) {

    # Extract Username and Password fields from the current account object
    $username = $account.Username
    $password = $account.Password

    # Create user account with the following settings: The user cannot change their password, and the password and account never expires.
    New-LocalUser -Name $username -password (ConvertTo-SecureString $password -AsPlainText -Force) -UserMayNotChangePassword -PasswordNeverExpires -AccountNeverExpires

    # Add the user to the Administrators group to grant it admin privileges
    Add-LocalGroupMember -Group "Administrators" -Member $username
}

# Removes all text from the current display
Clear-Host



<# Start Device Profiling Processes #>

# Change working directory to location of other files
Set-Location -Path $othersDir

# > Display device service tag
Write-Host "Service Tag: $serviceTag`n"

# > Display device network adapter information

# Query WMI Win32_NetworkAdapter class for relevant information, store it as a string in a variable, and output it
$netAdapterInfo = Get-WmiObject Win32_NetworkAdapter | Select-Object NetConnectionID, Name, MACAddress | Out-String
Write-Host "Network Adapters:`n$netAdapterInfo"

# > Display storage drive serial number

# Open CrystalDiskInfo utility
Start-Process -FilePath "CrystalDiskInfo8_12_0\DiskInfo64.exe"



# Pause to keep the console open
Write-Host "N1 Clean Manual Script 1 execution completed. Press any key to exit..."
[System.Console]::ReadKey() | Out-Null