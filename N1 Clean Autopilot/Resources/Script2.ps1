<# Initialize Global Variables and Functions #>

# Get the current folder path where installers are located
$installersDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Installers"

# Get root of system drive
$rootSysDriveDir = $env:SystemDrive

# Get the folder path for temporary local app data files
$locAppDataTmpDir = $env:Temp

# Get the current user's desktop folder path
$desktopPath = "$env:USERPROFILE\Desktop"

# Input GlobalProtect VPN software name
$globalprotectAppName = "GlobalProtect*"

# Input Microsoft Teams software name
$msTeamsAppName = "Microsoft Teams*"

# Get device manufacturer information
$deviceManufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer

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

# > Install HP Support Assistant

# Check that device manufacturer is HP before installing HP software update application
if ($deviceManufacturer -like "HP*") {

    Start-Process -FilePath "sp148716.exe" -ArgumentList /s, /a, /s
}

# > Install or repair GlobalProtect

# Query WMI Win32_Product class for applications installed via MSI to check if GlobalProtect is already installed, and install or repair it accordingly
$installCheck = $null -ne (Get-WmiObject Win32_Product | Where-Object { $_.Name -like $globalprotectAppName })

if ($installCheck) {

    Start-Process -FilePath msiexec.exe -ArgumentList "/fm GlobalProtect64-6.0.3.msi /passive" -Wait
} 

else {

    Start-Process -FilePath "GlobalProtect64-6.0.3.msi" -ArgumentList /passive -Wait
}

# > Install Microsoft Teams

# Check if Microsoft Teams is already installed and functioning properly by checking that its name appears in the list of start menu apps
$installCheck = $null -ne (Get-StartApps | Where-Object { $_.Name -like $msTeamsAppName })

# If Microsoft Teams is not installed, proceed with the installation process
if (-not $installCheck) {
    
    # Check that an active internet connection exists. Alert user and prevent installation process from proceeding otherwise
    if (-not (checkInternetConnection)) {

        Write-Host "Waiting for internet connection..."
    
        while (-not (checkInternetConnection)) {
    
            Start-Sleep -Seconds 4
        }
    
        Write-Host "Internet connection established. Proceeding with Teams installation..."
    }

    # Create a new instance of .NET object for sending and receiving data
    $webClient = New-Object System.Net.WebClient

    try {

        # Attempt to remove any existing Microsoft Teams applications and registry keys for all users, if they exist to ensure clean installation
        Get-AppxPackage -All $msTeamsAppName | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        reg delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Teams" /f >$null 2>&1

        # Input temporary directory path for storing Teams installer files
        $tempMsTeamsInstallerPath = Join-Path "$locAppDataTmpDir" "msTeamsInstaller"

        # Clean up existing folders before creating a new temporary directory for installer files
        Remove-Item -Path $tempMsTeamsInstallerPath -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -ItemType Directory -Path $tempMsTeamsInstallerPath -Force | Out-Null

        # Input URLs for downloading the Microsoft Teams installer files
        $msTeamsX64MsixDlLink = "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"
        $msTeamsBootStrapperDlLink = "https://statics.teams.cdn.office.net/production-teamsprovision/lkg/teamsbootstrapper.exe"

        # Download files using System.Net.WebClient class
        $webClient.DownloadFile($msTeamsX64MsixDlLink, "$tempMsTeamsInstallerPath\MSTeams-x64.msix")
        $webClient.DownloadFile($msTeamsBootStrapperDlLink, "$tempMsTeamsInstallerPath\teamsbootstrapper.exe")

        # Provision Teams app with installation files
        Start-Process -FilePath "$tempMsTeamsInstallerPath\teamsbootstrapper.exe" -ArgumentList "-p", "-o", "$tempMsTeamsInstallerPath\MSTeams-x64.msix" -Wait

        # Delete temporary folder for installer files
        Remove-Item -Path $tempMsTeamsInstallerPath -Recurse -Force -ErrorAction SilentlyContinue
    } 
    
    catch {

        # Log error message if installation process fails for any reason
        Write-Error "Microsoft Teams installation failed. Process will be skipped."
    }

    # Clean up and release resources
    $webClient.Dispose()
    $webClient = $null
} 



<# Start System Configuration Processes #>

# > Set Plugged In and Battery profiles to never turn off display or sleep

# User powercfg utility to modify power plan settings
powercfg /change monitor-timeout-ac 0
powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-dc 0
powercfg /change standby-timeout-dc 0

# > Disable Smart App Control in Windows Defender

# Path to Smart App Control settings in the registry
$smartAppControlKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"

# Set Smart App Control registry value to 0
Set-ItemProperty -Path $smartAppControlKeyPath -Name "VerifiedAndReputablePolicyState" -Value 0

# > Create shortcuts for both brand-specific software update apps and Windows Update in Desktop folder of currently logged-in user

# Check device manufacturer, and define specific application paths and shortcut names accordingly
$brandSwUpdateAppPath = $null
$brandSwUpdateAppShortcutName = $null

if ($deviceManufacturer -like "Dell*") {

    $brandSwUpdateAppPath = "$rootSysDriveDir\Program Files (x86)\Dell\CommandUpdate\DellCommandUpdate.exe"
    $brandSwUpdateAppShortcutName = "Dell Command Update.lnk"

}

elseif ($deviceManufacturer -like "LENOVO*") {

    $brandSwUpdateAppPath = "$rootSysDriveDir\Program Files (x86)\Lenovo\System Update\tvsu.exe"
    $brandSwUpdateAppShortcutName = "Lenovo System Update.lnk"
}

# Create a new WScript.Shell COM object
$wshShell = New-Object -ComObject WScript.Shell

# Create the shortcuts, and set their target paths and other necessary properties
if ($null -ne $brandSwUpdateAppPath) {

    $brandSwUpdateAppShortcutPath = Join-Path $desktopPath $brandSwUpdateAppShortcutName
    $brandSwUpdateAppShortcut = $wshShell.CreateShortcut($brandSwUpdateAppShortcutPath)
    $brandSwUpdateAppShortcut.TargetPath = $brandSwUpdateAppPath
    $brandSwUpdateAppShortcut.WorkingDirectory = (Split-Path -Path $brandSwUpdateAppPath)
    $brandSwUpdateAppShortcut.Save()
}

$winUpdateShortcutPath = Join-Path $desktopPath "Windows Update.lnk"
$winUpdateShortcut = $wshShell.CreateShortcut($winUpdateShortcutPath)
$winUpdateShortcut.TargetPath = "ms-settings:windowsupdate"
$winUpdateShortcut.Save()

# Clean up and release resources
$wshShell = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# Removes all text from the current display
Clear-Host



# Pause to keep the console open
Write-Host "N1 Clean Autopilot Script 2 execution completed. Press any key to exit..."
[System.Console]::ReadKey() | Out-Null