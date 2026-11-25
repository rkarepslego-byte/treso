@echo off
rem Lance Installer.ps1 sans modifier la strategie d execution PowerShell du poste
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Installer.ps1" %*
pause
