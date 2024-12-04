<# Initialize Global Variables and Functions #>

# Get the current folder path where installers are located
$installersDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Installers"

# Get the folder path for temporary local app data files
$locAppDataTmpDir = $env:Temp

# Get the current user's desktop folder path
$desktopPath = [Environment]::GetFolderPath("Desktop")

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

# > Install HP Support Assistant

# Check that device manufacturer is HP before installing HP software update application
if ($deviceManufacturer -like "HP*") {

    Start-Process -FilePath "sp148716.exe" -ArgumentList /s, /a, /s
}

# > Install or repair GlobalProtect

# Query Win32_Product WMI class to check if GlobalProtect is already installed, and install or repair it accordingly
$installCheck = Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -like "$globalprotectAppName" }

if ($installCheck) {

    Start-Process -FilePath msiexec.exe -ArgumentList "/fm GlobalProtect64-6.0.3.msi /passive" -Wait
} 

else {

    Start-Process -FilePath GlobalProtect64-6.0.3.msi -ArgumentList /passive -Wait
}

# > Install Microsoft Teams

$installCheck = Get-StartApps | Where-Object { $_.Name -like "$msTeamsAppName" }

if (-not $installCheck) {
    
    if (-not (checkInternetConnection)) {

        Write-Host "Waiting for internet connection..."
    
        while (-not (checkInternetConnection)) {
    
            Start-Sleep -Seconds 4
        }
    
        Write-Host "Internet connection established. Continuing with installation...`n"
    }

    try {

        Get-AppxPackage -all $msTeamsAppName | Remove-AppxPackage -allusers
        reg delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Teams" /f >$null 2>&1

        $tempMsTeamsInstallerPath = Join-Path "$locAppDataTmpDir" "msTeamsInstaller"

        # Check for and clean up temporary folder for installer files
        if (Test-Path $tempMsTeamsInstallerPath) {

            Remove-Item -Path $tempMsTeamsInstallerPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $tempMsTeamsInstallerPath -Force

        $msTeamsX64MsixDlLink = "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"
        $msTeamsBootStrapperDlLink = "https://statics.teams.cdn.office.net/production-teamsprovision/lkg/teamsbootstrapper.exe"

        Start-BitsTransfer -Source $msTeamsX64MsixDlLink -Destination "$tempMsTeamsInstallerPath\MSTeams-x64.msix"
        Start-BitsTransfer -Source $msTeamsBootStrapperDlLink -Destination "$tempMsTeamsInstallerPath\teamsbootstrapper.exe"

        & "$tempMsTeamsInstallerPath\teamsbootstrapper.exe" -p -o "$tempMsTeamsInstallerPath\MSTeams-x64.msix" >$null 2>&1

        # Delete temporary folder for installer files
        Remove-Item -Path $tempMsTeamsInstallerPath -Recurse -Force
    } 
    
    catch {

        Write-Error "Microsoft Teams installation failed. Process will be skipped."
    }
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

    $brandSwUpdateAppPath = "C:\Program Files (x86)\Dell\CommandUpdate\DellCommandUpdate.exe"
    $brandSwUpdateAppShortcutName = "Dell Command Update.lnk"

}

elseif ($deviceManufacturer -like "LENOVO*") {

    $brandSwUpdateAppPath = "C:\Program Files (x86)\Lenovo\System Update\tvsu.exe"
    $brandSwUpdateAppShortcutName = "Lenovo System Update.lnk"
}

# Define the full path for the shortcuts
$brandSwUpdateAppShortcutPath = Join-Path $desktopPath $brandSwUpdateAppShortcutName

$winUpdateShortcutPath = Join-Path $desktopPath 'Windows Update.lnk'

# Create a new WScript.Shell COM object
$wshShell = New-Object -ComObject WScript.Shell

# Create the shortcuts
$brandSwUpdateAppShortcut = $wshShell.CreateShortcut($brandSwUpdateAppShortcutPath)

$winUpdateShortcut = $wshShell.CreateShortcut($winUpdateShortcutPath)

# Set the target paths and other shortcut properties
$brandSwUpdateAppShortcut.TargetPath = $brandSwUpdateAppPath
$brandSwUpdateAppShortcut.WorkingDirectory = (Split-Path -Path $brandSwUpdateAppPath)
$brandSwUpdateAppShortcut.Save()

$winUpdateShortcut.TargetPath = 'ms-settings:windowsupdate'
$winUpdateShortcut.Save()

<#
# Pause to keep the console open
Write-Host "`nPress any key to exit..."
[System.Console]::ReadKey() | Out-Null
#>