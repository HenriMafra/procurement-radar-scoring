@echo off
title ENTERPRISECORE Radar - Publicar nova versao no GitHub
cd /d "%~dp0"
echo.
echo  Publica o estado atual no GitHub como uma nova versao.
echo  Dados reais (boletins, contatos, planilha) NAO sobem.
echo.
set /p MSG="Descreva o que mudou (Enter para o padrao): "
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\publicar-versao.ps1" -Mensagem "%MSG%"
pause
