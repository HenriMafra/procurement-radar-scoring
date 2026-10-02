@echo off
title ENTERPRISECORE Radar - Preparar copia para entrega
cd /d "%~dp0"
echo.
echo Este script cria uma COPIA limpa do projeto numa pasta nova,
echo sem os seus dados pessoais. A sua instalacao NAO e alterada.
echo.
pause
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\preparar-entrega.ps1"
pause
