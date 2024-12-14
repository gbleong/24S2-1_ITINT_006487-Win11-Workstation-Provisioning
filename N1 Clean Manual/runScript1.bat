:: Instructions: Run this once after logging in to placeholder Windows local admin user account

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Resources\Script1.ps1""' -Verb RunAs -Wait"