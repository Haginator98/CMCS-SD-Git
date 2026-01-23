@echo off
REM CMCS Servicedesk Tools Launcher
REM Double-click this file to start the Servicedesk Tools
REM Made by Mr. Hagen - 2026

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Tools.ps1"

pause
