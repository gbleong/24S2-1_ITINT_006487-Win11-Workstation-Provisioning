:: Instructions: Run this once, before starting the Autopilot process in the Out-of-Box Experience (OOBE) setup environment

powershell -NoProfile -Command "Get-ChildItem -Path '%~dp0' -Recurse | ForEach-Object { Unblock-File -Path $_.FullName }; Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0\Resources\Script1.ps1\"' -Verb RunAs -Wait"