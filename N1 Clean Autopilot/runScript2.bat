:: Instructions: Run this once, only after logging in to any Windows account with administrator rights

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Resources\Script2.ps1""' -Verb RunAs -Wait"