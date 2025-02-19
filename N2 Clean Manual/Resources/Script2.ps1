<# Initialize Global Variables and Functions #>

# Get the current folder path where credential files are located
$credentialFilesDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Credential Files"

# Get device service tag from Win32_BIOS WMI class
$serviceTag = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber

# Get Windows version and build
$winVerBuild = "$((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion) (OS Build $((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuildNumber).$((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').UBR))"



<# Start System Configuration Processes #>

# Change working directory to location of credential files
Set-Location -Path $credentialFilesDir

# > Delete redundant Windows user accounts and files

# Import list of allowed usernames from CSV files containing Windows account credentials as an array of objects where each row represents an account
$n2BreakGlassAdminAccounts = Import-Csv -Path "N2 Windows Break-Glass Admin Credentials.csv" | Select-Object -ExpandProperty Username
$n2EmployeeAccounts = Import-Csv -Path "N2 Windows Employee Admin & User Credentials.csv" | Select-Object -ExpandProperty Username
$n2Accounts = $n2BreakGlassAdminAccounts + $n2EmployeeAccounts

# Get list of all local user accounts excluding system default accounts
$existingLocalAccounts = Get-LocalUser | Where-Object { $_.Name -notmatch "^DefaultAccount|^WDAGUtilityAccount|^Guest|^Administrator" }

# Iterate through each account object in the imported CSV
foreach ($account in $existingLocalAccounts) {

    $username = $account.Name

    # Executes subsequent code if username is not in the list of allowed accounts
    if ($username -notin $n2Accounts) {

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

# Removes all text from the current display
Clear-Host



<# Start Device Profiling Processes #>

# > Display device service tag
Write-Host "Service Tag: $serviceTag`n"

# > Open Windows version and build

Write-Host "Windows Version and Build: $winVerBuild`n"

# > Display device network adapter information

# Query WMI Win32_NetworkAdapter class for relevant information, store it as a string in a variable, and output it
$netAdapterInfo = Get-CimInstance -ClassName Win32_NetworkAdapter | Select-Object NetConnectionID, Name, MACAddress | Out-String
Write-Host "Network Adapters:`n$netAdapterInfo"

# > Display storage drive serial number

# Change working directory to location of other files
Set-Location -Path $othersDir

# Open CrystalDiskInfo utility
Start-Process -FilePath "CrystalDiskInfo8_12_0\DiskInfo64.exe"

# > Display computer name

Start-Process sysdm.cpl



# Pause to keep the console open
Write-Host "N2 Clean Manual Script 2 execution completed. Press any key to exit..."
[System.Console]::ReadKey() | Out-Null