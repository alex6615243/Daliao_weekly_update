@echo off
chcp 65001 >nul
title 大寮召會週訊一鍵更新 2026.08.15-v5
echo 啟動位置：%~dp0
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0一鍵更新大寮周訊_桌面版.ps1"
echo.
pause
