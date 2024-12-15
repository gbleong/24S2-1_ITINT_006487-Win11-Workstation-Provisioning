<# Initialize Global Variables and Functions #>

# Get the folder path for temporary local app data files
$locAppDataTmpDir = $env:Temp

# Get the current user's desktop folder path
$desktopPath = "$env:USERPROFILE\Desktop"

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

# > Install Microsoft Teams

# Check if Microsoft Teams is already installed and functioning properly by checking that its name appears in the list of start menu apps
$installCheck = Get-StartApps | Where-Object { $_.Name -like "$msTeamsAppName" }

# If Microsoft Teams is not installed, proceed with the installation process
if (-not $installCheck) {
    
    # Check that an active internet connection exists. Alert user and prevent installation process from proceeding otherwise
    if (-not (checkInternetConnection)) {

        Write-Host "Waiting for internet connection..."
    
        while (-not (checkInternetConnection)) {
    
            Start-Sleep -Seconds 4
        }
    
        Write-Host "Internet connection established. Continuing with installation...`n"
    }

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
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($msTeamsX64MsixDlLink, "$tempMsTeamsInstallerPath\MSTeams-x64.msix")
        $webClient.DownloadFile($msTeamsBootStrapperDlLink, "$tempMsTeamsInstallerPath\teamsbootstrapper.exe")
        $webClient.Dispose()
        $webClient = $null

        # Provision Teams app with installation files
        Start-Process -FilePath "$tempMsTeamsInstallerPath\teamsbootstrapper.exe" -ArgumentList "-p", "-o", "$tempMsTeamsInstallerPath\MSTeams-x64.msix" -Wait

        # Delete temporary folder for installer files
        Remove-Item -Path $tempMsTeamsInstallerPath -Recurse -Force -ErrorAction SilentlyContinue
    } 
    
    catch {

        # Log error message if installation process fails for any reason
        Write-Error "Microsoft Teams installation failed. Process will be skipped."
    }
} 



<# Start System Configuration Processes #>
<#
# Change working directory to location of credential files
Set-Location -Path $credentialFilesDir

# > Delete redundant Windows user accounts WIP

$n1WinLocalAdminCredFile = "N1 Windows Local Admin Credentials.csv"

if (Test-Path $n1WinLocalAdminCredFile) {

    # Import list of allowed usernames from the CSV file containing credentials for N1 Windows local admin accounts
    $n1AdminAccounts = Import-Csv -Path $n1WinLocalAdminCredFile | Select-Object -ExpandProperty Username

    # Get a list of all local user accounts excluding system default accounts
    $existingLocalAccounts = Get-LocalUser | Where-Object { $_.Name -notmatch "^DefaultAccount|^WDAGUtilityAccount|^Guest|^Administrator|^crdsecagent\$admin$" }

    foreach ($account in $existingLocalAccounts) {

        $Username = $account.Username

        # Check if the username is not in the list of allowed accounts
        if ($Username -notin $n1AdminAccounts) {

            Write-Host "Deleting user account: $Username"

            try {

                # Remove the user account
                Remove-LocalUser -Name $Username
                Write-Host "Successfully deleted user account: $Username"
            } 
            
            catch {

                Write-Host "Failed to delete user account: $Username. Error: $_"
            }
        } 
        
        else {
            Write-Host "Skipping allowed user account: $Username"
        }
    }
} 

else {

    Write-Host "Windows local user account clean up failed. Process will be skipped."
}
#>
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

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wshShell) | Out-Null
$wshShell = $null

<#
# Pause to keep the console open
Write-Host "`nPress any key to exit..."
[System.Console]::ReadKey() | Out-Null
#>