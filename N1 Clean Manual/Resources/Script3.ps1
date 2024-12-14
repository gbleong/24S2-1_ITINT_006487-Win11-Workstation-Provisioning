Start-Process -FilePath msiexec.exe -ArgumentList "/i xagtSetup_35.31.28_universal.msi /passive" -Wait

Start-Process -FilePath GlobalProtect64-6.0.3.msi -ArgumentList /passive -Wait

Start-Process -FilePath msiexec.exe -ArgumentList "/i ClearPassOnGuardInstall.msi /passive" -Wait

Start-Process -FilePath ".\Tanium 7.4.10.1075\SetupClient.exe" -ArgumentList /S

#winver and sysdm