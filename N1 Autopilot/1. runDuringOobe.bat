@echo off

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Software Installation\installOobeViableSoftware.ps1""' -Verb RunAs -Wait -WindowStyle Hidden"

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Device Profiling\listDeviceInfo.ps1""' -Verb RunAs -Wait"