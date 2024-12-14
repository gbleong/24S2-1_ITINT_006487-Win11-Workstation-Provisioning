:: Instructions: Run this once after joining the device to the organizational Windows Active Directory domain

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Resources\Script3.ps1""' -Verb RunAs -Wait"