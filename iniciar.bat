@echo off
chcp 65001 > nul
title Automacao Power BI + Excel 
echo ================================================
echo   AUTOMACAO POWER BI + EXCEL 
echo ================================================
echo.
echo Iniciando monitoramento...
echo Log visual: abra a pasta 'logs' e clique em historico.html
echo (Pressione Ctrl+C para encerrar)
echo.
cd /d "%~dp0"
python automacao_powerbi.py
pause
