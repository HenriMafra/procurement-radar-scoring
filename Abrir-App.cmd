@echo off
title ENTERPRISECORE Radar
cd /d "%~dp0"
echo ============================================================
echo   ENTERPRISECORE Radar - Licitacoes
echo   Iniciando o servidor local (Node)...
echo ============================================================

REM Sobe o servidor numa janela propria (nao feche enquanto usar o app)
start "ENTERPRISECORE Radar (servidor - NAO FECHE)" node "%~dp0app\server.js"

echo Aguardando o servidor responder em http://localhost:8790 ...
:wait
timeout /t 1 >nul
powershell -NoProfile -Command "try{ $null=(Invoke-WebRequest http://localhost:8790 -UseBasicParsing -TimeoutSec 1); exit 0 }catch{ exit 1 }"
if errorlevel 1 goto wait

start "" http://localhost:8790
echo.
echo App aberto no navegador. Pode minimizar esta janela.
echo Para ENCERRAR: feche a janela "ENTERPRISECORE Radar (servidor - NAO FECHE)".
timeout /t 4 >nul
