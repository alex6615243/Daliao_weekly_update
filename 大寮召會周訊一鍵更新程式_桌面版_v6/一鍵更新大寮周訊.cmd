@echo off
chcp 65001 >nul
title Weekly Bulletin Updater 2026.08.15-v6
echo Program folder: %~dp0
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0updater.ps1"
echo.
pause
