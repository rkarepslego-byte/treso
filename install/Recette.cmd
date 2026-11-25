@echo off
rem Lance Recette.ps1 sans modifier la strategie d execution PowerShell du poste
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Recette.ps1" %*
pause
