<# Initialize Global Variables and Functions #>

# Get the current folder path where installers are located
$installersDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Installers"



<# Start Software Installation Processes #>

# Change working directory to location of installers
Set-Location -Path $installersDir

# > Install Tanium

Start-Process -FilePath "Tanium 7.4.10.1075\SetupClient.exe" -ArgumentList /S

# > Install Trellix Endpoint Security Agent

Start-Process -FilePath msiexec.exe -ArgumentList "/i IMAGE_HX_AGENT_WIN_35.31.28\xagtSetup_35.31.28_universal.msi /passive" -Wait

# > Install GlobalProtect

Start-Process -FilePath GlobalProtect64-6.0.3.msi -ArgumentList /passive -Wait

# > Install ClearPass OnGuard

Start-Process -FilePath msiexec.exe -ArgumentList "/i ClearPassOnGuardInstall.msi /passive" -Wait



<# Start Device Profiling Processes #>

# > Open About Windows dialog

Start-Process winver.exe

# > Open System Properties dialog

Start-Process sysdm.cpl

<#
# Pause to keep the console open
Write-Host "`nPress any key to exit..."
[System.Console]::ReadKey() | Out-Null
#>