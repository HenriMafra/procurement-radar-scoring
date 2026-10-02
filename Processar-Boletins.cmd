@echo off
title ENTERPRISECORE Radar - Processar Boletins
cd /d "%~dp0"
echo ============================================================
echo   ENTERPRISECORE Radar - Processamento dos boletins da Conlicitacao
echo ------------------------------------------------------------
echo   Coloque os .xlsx dos boletins em:  data\boletins
echo   (pode ser 1 ou os 3 de uma vez)
echo ============================================================
echo.
REM -Zerar = o resultado reflete SO os boletins que estao na pasta agora.
REM (mesmo comportamento do botao "Processar boletins" do app; remova o -Zerar para acumular historico)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\processar-boletins.ps1" -Zerar
echo.
echo Pronto. A planilha esta em  data\Planilha-ENTERPRISECORE-Licitacoes.xlsx
echo e o app ja passa a mostrar estes dados (data\dados.json).
echo.
pause
