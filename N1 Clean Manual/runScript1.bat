:: Instructions: Run once after logging in to placeholder Windows account

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0\Resources\Script1.ps1""' -Verb RunAs -Wait"