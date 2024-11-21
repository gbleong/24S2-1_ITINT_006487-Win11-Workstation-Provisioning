@echo off
:: Request administrative privileges
powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0runInstPkgs.ps1""' -Verb RunAs -WindowStyle Hidden"