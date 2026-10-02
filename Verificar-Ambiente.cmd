@echo off
title ENTERPRISECORE Radar - Verificar Ambiente
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\verificar-ambiente.ps1"
echo.
pause
