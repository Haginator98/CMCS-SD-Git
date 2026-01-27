@echo off

cd /d "%~dp0"
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Tools.ps1"

pause
