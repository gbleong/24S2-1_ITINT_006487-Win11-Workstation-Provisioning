This repository is designed to streamline and automate the process of configuring and setting up PC assets owned by an anonymous organization, within an environment requiring consistent and repeatable configurations. By utilizing a structured file system and a combination of batch and PowerShell scripts, this project ensures efficient setup while maintaining flexibility for different PC security classifications.

=== Features ===

1. Software Installation: Automated installation of required applications.

2. System Configuration: Perform necessary configurations to the operating system without manual intervention.

3. Device Profiling: Employs various utilities to retrieve and display essential system information.

=== File Structure Overview ===

Folder Structure:

```plaintext
Folders Labelled with Security Classification and Setup Method
    └─ Batch scripts (to run PowerShell scripts with admin privileges)
    └─ Resources Folder
        └─ PowerShell Scripts
        └─ Installers
        └─ Utility Tools
        └─ Credential Files (CSV files containing user credentials)
```

Key Components:

1. Folders Labelled with Security Classification and Setup Method

    - Each folder is given the appropriate label corresponding to a specific PC security classification and method of configuration to be employed.

    - Users should use files from the correct folder to ensure the proper setup process is followed accordingly.

2. Batch Scripts

    - These scripts act as entry points, ensuring that the PowerShell scripts in the Resources folder are executed with elevated privileges and bypass execution policy restrictions.

3. PowerShell Scripts

    - Contains the main code to automate the system setup tasks, which also makes use of various executable files and applications located within the Resources folder.

4. Installers Folder

    - Houses all installer executables and packages for the deployment of required software.

5. Utility Tools Folder

    - Contains additional tools for gathering detailed device information.

6. Credential Files Folder

    - Stores various credentials in CSV format.

=== Usage Instructions ===

Prerequisites:

1. Administrator Access

    - Ensure the user is utilizing an account with administrative privileges on the target machine.

    - Some scripts require elevated permissions to modify system settings.

2. Execution Policy

    - The scripts are designed to bypass execution policy restrictions; however, verify that the system allows PowerShell scripts to execute.

Steps to Execute:

1. Clone or download the repository to the target machine.

2. Navigate to the folder corresponding to the specific PC security classification and method of configuration to be employed.

3. Move all required files such as installer executables, packages, and other tools into their appropriate folders. Refer to the file paths mentioned in the PowerShell scripts if unclear.

4. Follow instructions provided within batch scripts for more directions on when to run them.