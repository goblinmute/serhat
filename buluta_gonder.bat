@echo off
title Arbitraj Botu - Buluta Gonderiliyor...
echo ========================================
echo   BOT GUNCELLEME VE DEPLOY BASLIYOR
echo ========================================
echo.

:: PowerShell scriptini calistir
powershell -ExecutionPolicy Bypass -File "%~dp0deploy.ps1"

echo.
echo ========================================
echo   ISLEM TAMAMLANDI!
echo ========================================
pause
