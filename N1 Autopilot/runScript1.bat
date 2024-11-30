:: Instructions: Run this once, when in the Out-of-Box Experience (OOBE) setup environment after installing Windows 11

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Resources\Script1.ps1""' -Verb RunAs -Wait"