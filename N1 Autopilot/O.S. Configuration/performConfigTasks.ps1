<# Start Administrative Processes #>

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