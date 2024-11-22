<# Initialize Variables #>

# Get the current folder path where the script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Get the folder path for temporary local app data files
$locAppDataTmpDir = Join-Path $env:LOCALAPPDATA "Temp"

# Get device manufacturer information
$deviceManufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer

<# Start installer processes #>

# Change working directory to script directory
Set-Location -Path $scriptDir

# > Intel Bluetooth Driver

Start-Process -FilePath "BT-23.60.0-64UWD-Win10-Win11.exe" -ArgumentList /passive

# > Intel Wi-Fi Driver

Start-Process -FilePath "WiFi-23.60.1-Driver64-Win10-Win11.exe" -ArgumentList /passive

# > Intel Network Connections Driver

# Get list of folders in directory for temporary local app data files
$locAppDataTmpFolders = Get-ChildItem -Path $locAppDataTmpDir -Directory -ErrorAction SilentlyContinue

# Check for and remove leftover temporary WinZip Self-Extractor files
foreach ($folder in $locAppDataTmpFolders) {

    if ($folder.Name -match "^WZSE\d+\.TMP$") {

        Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
    }
}

Start-Process -FilePath "Wired_driver_29.1_x64.exe"

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

# > Comber

Start-Process -FilePath "Comber x64 1.0.1.0.exe" -ArgumentList /v"/passive" -Wait

# > Dell Command Update

# Check that device manufacturer is Dell before installing Dell-specific software update application
if ($deviceManufacturer -like "Dell*") {

    Start-Process -FilePath "Dell-Command-Update-Application_T45GH_WIN_5.3.0_A00.exe" -ArgumentList /passthrough, /v"/passive" -Wait
}

# > Lenovo System Update

# Check that device manufacturer is Lenovo before installing Lenovo-specific software update application
if ($deviceManufacturer -like "LENOVO*") {

    Start-Process -FilePath "system_update_5.08.02.25.exe" -ArgumentList /SP-, /SILENT -Wait
}

<#
# Pause to keep the console open
Write-Host "`nPress any key to exit..."
[System.Console]::ReadKey() | Out-Null
#>