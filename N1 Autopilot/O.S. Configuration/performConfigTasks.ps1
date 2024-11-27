<# Initialize Global Variables #>

# Get the current user's desktop folder path
$desktopPath = [Environment]::GetFolderPath("Desktop")

# Get device manufacturer information
$deviceManufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer

<# Start Configuration Processes #>

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
$brandSwUpdateAppShortcutPath = Join-Path -Path $desktopPath -ChildPath $brandSwUpdateAppShortcutName

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