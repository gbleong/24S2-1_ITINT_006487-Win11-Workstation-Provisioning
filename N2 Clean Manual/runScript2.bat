:: Instructions: Run this once, only after logging in to a N2 Windows break-glass or employee admin account

powershell -NoProfile -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Resources\Script2.ps1""' -Verb RunAs -Wait"