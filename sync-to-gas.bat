@echo off
chcp 65001 >nul
powershell -ExecutionPolicy Bypass -File "%~dp0sync-to-gas.ps1"
pause
