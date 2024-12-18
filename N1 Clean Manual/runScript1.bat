:: Instructions: Run this once after logging in to placeholder Windows local admin user account

powershell -NoProfile -Command "Get-ChildItem -Path '%~dp0' -Recurse | ForEach-Object { Unblock-File -Path $_.FullName }; Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0\Resources\Script1.ps1\"' -Verb RunAs -Wait"