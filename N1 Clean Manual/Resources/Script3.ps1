<# Initialize Global Variables and Functions #>

# Get the current folder path where installers are located
$installersDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Installers"

# Get computer name
$compName = (Get-WmiObject Win32_ComputerSystem).Name

# Get Windows version and build
$winVerBuild = "$((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion) (OS Build $((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuildNumber).$((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').UBR))"



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

# Removes all text from the current display
Clear-Host


<# Start Device Profiling Processes #>

# > Display computer name

Write-Host "Computer Name: $compName`n"

# > Open Windows version and build

Write-Host "Windows Version and Build: $winVerBuild`n"



# Pause to keep the console open
Write-Host "N1 Clean Manual Script 3 execution completed. Press any key to exit..."
[System.Console]::ReadKey() | Out-Null