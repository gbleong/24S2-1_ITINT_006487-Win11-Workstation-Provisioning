@echo off

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Software Installation\installRemainingSoftware.ps1""' -Verb RunAs -Wait -WindowStyle Hidden"

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\O.S. Configuration\performConfigTasks.ps1""' -Verb RunAs -Wait -WindowStyle Hidden"