:: Instructions: Run this once, before starting the Autopilot process in the Out-of-Box Experience (OOBE) setup environment

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Resources\Script1.ps1""' -Verb RunAs -Wait"