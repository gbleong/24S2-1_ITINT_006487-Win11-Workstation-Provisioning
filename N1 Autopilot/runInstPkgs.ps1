# Get the current folder path where the script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define the directory to monitor
$locAppDataTmpDir = Join-Path $env:LOCALAPPDATA "Temp"

# Change working directory to script directory
Set-Location -Path $scriptDir

# Start installer processes

# Intel Bluetooth Drivers
#Start-Process -FilePath "BT-23.60.0-64UWD-Win10-Win11.exe" -ArgumentList /passive -Wait

# Intel Wi-Fi Drivers
#Start-Process -FilePath "WiFi-23.60.1-Driver64-Win10-Win11.exe" -ArgumentList /passive -Wait

# Get the list of folders in the Temp directory
$locAppDataTmpDirFolders = Get-ChildItem -Path $locAppDataTmpDir -Directory -ErrorAction SilentlyContinue

# Check if any folder matches the naming convention
foreach ($folder in $locAppDataTmpDirFolders) {

    if ($folder.Name -match "^WZSE\d+\.TMP$") {

        Remove-Item -Path $folder.Name -Recurse -Force -ErrorAction Stop
    }
}

# Intel Network Connections Software
Start-Process -FilePath "Wired_driver_29.1_x64.exe"

$targetWzseTmpFolder = $null

do {

    $locAppDataTmpFolders = Get-ChildItem -Path $locAppDataTmpDir -Directory -ErrorAction SilentlyContinue

    # Check if any folder matches the naming convention
    foreach ($folder in $locAppDataTmpFolders) {

        if ($folder.Name -match "^WZSE\d+\.TMP$") {

            $targetWzseTmpFolder = $folder.Name

	        do {
    		    # Construct the expected file path
                $setupbdExe = Join-Path $locAppDataTmpDir (Join-Path $targetWzseTmpFolder "APPS\SETUP\SETUPBD\Winx64\SetupBD.exe")
                Write-Host $setupbdExe

                # Check if the file exists
                if (Test-Path -Path $setupbdExe) {

	                Start-Process -FilePath $setupbdExe -ArgumentList "/switch /m" -Wait
                    break
                }

                # Pause for a short time to reduce CPU usage
                Start-Sleep -Seconds 2
            } while (-not (Test-Path -Path $setupbdExe))

            break
        }
    }

    # Pause for a short time to reduce CPU usage
    Start-Sleep -Seconds 2
} while (-not $targetWzseTmpFolder)

Stop-Process -Name "SetupBD"

#Start-Process -FilePath "Comber x64 1.0.1.0.exe" -ArgumentList /v"/passive" -Wait

# Pause to keep the console open
#Write-Host "`nPress any key to exit..."
#[System.Console]::ReadKey() | Out-Null