@echo off

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Software Installation\installOobeViableSoftware.ps1""' -Verb RunAs -WindowStyle Hidden"

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\DeviceProfiling\listSt&Mac&SdSn.ps1""' -Verb RunAs -WindowStyle Hidden"