<# Initialize Global Variables and Functions #>

# Get the current folder path where credential files are located
$credentialFilesDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Credential Files"

# Get root of system drive
$rootSysDriveDir = $env:SystemDrive

# Get the folder path for temporary local app data files
$locAppDataTmpDir = $env:Temp

# Get the current user's desktop folder path
$desktopPath = "$env:USERPROFILE\Desktop"

# Input Microsoft Teams software name
$msTeamsAppName = "Microsoft Teams*"

# Get device manufacturer information
$deviceManufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer

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

    try {

        # Attempt to remove any existing Microsoft Teams applications and registry keys for all users, if they exist to ensure clean installation
        Get-AppxPackage -All $msTeamsAppName | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        reg.exe delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Teams" /f >$null 2>&1

        # Input temporary directory path for storing Teams installer files
        $tempMsTeamsInstallerPath = Join-Path "$locAppDataTmpDir" "msTeamsInstaller"

        # Clean up existing folders before creating a new temporary directory for installer files
        Remove-Item -Path $tempMsTeamsInstallerPath -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -ItemType Directory -Path $tempMsTeamsInstallerPath -Force | Out-Null

        # Input URLs for downloading the Microsoft Teams installer files
        $msTeamsX64MsixDlLink = "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"
        $msTeamsBootStrapperDlLink = "https://statics.teams.cdn.office.net/production-teamsprovision/lkg/teamsbootstrapper.exe"

         # Download files using Background Intelligent Transfer Service (BITS)
        Start-BitsTransfer -Source $msTeamsX64MsixDlLink -Destination "$tempMsTeamsInstallerPath\MSTeams-x64.msix"
        Start-BitsTransfer -Source $msTeamsBootStrapperDlLink -Destination "$tempMsTeamsInstallerPath\teamsbootstrapper.exe"

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

# Change working directory to location of credential files
Set-Location -Path $credentialFilesDir

# > Delete redundant Windows user accounts and files

# Import list of allowed usernames from CSV file containing Windows local admin credentials as an array of objects where each row represents an account
$n1AdminAccounts = Import-Csv -Path "N1 Windows Local Admin Credentials.csv" | Select-Object -ExpandProperty Username

# Get list of all local user accounts excluding system default accounts
$existingLocalAccounts = Get-LocalUser | Where-Object { $_.Name -notmatch "^DefaultAccount|^WDAGUtilityAccount|^Guest|^Administrator" }

# Iterate through each account object in the imported CSV
foreach ($account in $existingLocalAccounts) {

    $username = $account.Name

    # Executes subsequent code if username is not in the list of allowed accounts
    if ($username -notin $n1AdminAccounts) {

        # Define required native API calls and structures
        Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    // Define class for interacting with wtsapi32.dll
    public class WtsApi32 {

        // Define constant for querying username
        public const int WTSUserName = 5;

        // Define structure representing session information returned by WTSEnumerateSessions
        [StructLayout(LayoutKind.Sequential)]
        public struct WTS_SESSION_INFO {

            public int SessionID;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pWinStationName;
            public int State;
        }

        // Define function to enumerate active sessions
        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSEnumerateSessions(

            IntPtr hServer,
            int Reserved,
            int Version,
            out IntPtr ppSessionInfo,
            out int pCount
        );

        // Define function to query user information for sessions
        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSQuerySessionInformation(

            IntPtr hServer,
            int SessionID,
            int WTSInfoClass,
            out IntPtr ppBuffer,
            out int pBytesReturned
        );

        // Define function to free allocated memory
        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern void WTSFreeMemory(IntPtr pMemory);

        // Define function to log off session
        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSLogoffSession(

            IntPtr hServer,
            int SessionID,
            bool bWait
        );
    }
"@

        # Define variable for using local server handle
        $serverHandle = [IntPtr]::Zero

        # Define variable for holding session information pointer
        $sessionInfoPtr = [IntPtr]::Zero

        # Define variable to store number of active sessions
        $sessionCount = 0

        # Call WTSEnumerateSessions to get list of all user sessions
        [WtsApi32]::WTSEnumerateSessions($serverHandle, 0, 1, [ref]$sessionInfoPtr, [ref]$sessionCount)

        # Determine size of each session structure to correctly parse memory block
        $structSize = [System.Runtime.InteropServices.Marshal]::SizeOf([System.Type] ([WtsApi32+WTS_SESSION_INFO]))
    
        # Iterate through retrieved sessions
        for ($i = 0; $i -lt $sessionCount; $i++) {
    
            # Calculate pointer to the current session information structure
            $currentSessionPtr = [IntPtr]::Add($sessionInfoPtr, $i * $structSize)
    
            # Get session information
            $session = [System.Runtime.InteropServices.Marshal]::PtrToStructure($currentSessionPtr, [System.Type] ([WtsApi32+WTS_SESSION_INFO]))
    
            # Query session for username using WTSQuerySessionInformation
            $buffer = [IntPtr]::Zero
            $bytesReturned = 0

            if ([WtsApi32]::WTSQuerySessionInformation($serverHandle, $session.SessionID, [WtsApi32]::WTSUserName, [ref]$buffer, [ref]$bytesReturned)) {

                # Convert retrieved username from ANSI string to readable PowerShell string
                $sessionUsername = [System.Runtime.InteropServices.Marshal]::PtrToStringAnsi($buffer)

                # Free up resources
                [WtsApi32]::WTSFreeMemory($buffer)
        
                # Check if retrieved username matches target username, and log off matching user
                if ($sessionUsername -and ($sessionUsername -ieq $username)) {

                    [WtsApi32]::WTSLogoffSession($serverHandle, $session.SessionID, $true)
                }
            }
        }

        # Free memory allocated for session list
        [WtsApi32]::WTSFreeMemory($sessionInfoPtr)

        # Get active processes for the user
        $userActiveProcesses = Get-Process -IncludeUserName | Where-Object { $_.UserName -like "*\$username" }

        if ($userActiveProcesses) {

            # Iterate through and terminate each process
            foreach ($process in $userActiveProcesses) {

                Stop-Process -Id $process.Id -Force -ErrorAction -SilentlyContinue
            }
        }

        # Remove the user account
        Remove-LocalUser -Name $username

        # Grants full access permissions to Administrators group using  Integrity Control Access Control Lists (ICACLS) utility to clean up leftover user profile files
        icacls.exe "$rootSysDriveDir\Users\$username" /grant Administrators:F /T /C >$null 2>&1
        Remove-Item -Path "$rootSysDriveDir\Users\$username" -Recurse -Force -ErrorAction SilentlyContinue
    } 
}

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

# Removes all text from the current display
Clear-Host



# Pause to keep the console open
Write-Host "N1 Clean Manual Script 2 execution completed. Press any key to exit..."
[System.Console]::ReadKey() | Out-Null