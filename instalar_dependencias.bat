@echo off
chcp 65001 > nul
echo ================================================
echo   INSTALACAO DE DEPENDENCIAS - Auto Power BI
echo ================================================
echo.

echo [1/5] Verificando Python...
python --version
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado!
    echo Baixe em: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo [2/5] Instalando dependencias principais...
pip install openpyxl watchdog schedule --upgrade

echo.
echo [3/5] Instalando pyautogui (clique automatico)...
pip install pyautogui pygetwindow --upgrade

echo.
echo [4/5] Verificando instalacao...
python -c "import openpyxl, watchdog, schedule, pyautogui, pygetwindow; print('[OK] Todas as dependencias instaladas!')"
if %errorlevel% neq 0 (
    echo [AVISO] pyautogui pode ter falhado - o restante da automacao funciona normalmente.
    echo         Apenas o clique automatico ficara desativado.
)

echo.
echo [5/5] Criando pastas...
if not exist "logs" mkdir logs
if not exist "backups" mkdir backups
echo [OK] Pastas criadas.

echo.
echo ================================================
echo   INSTALACAO CONCLUIDA!
echo ================================================
echo.
echo Proximos passos:
echo  1. Edite o config.json com seus caminhos
echo  2. Execute iniciar.bat para rodar
echo  3. Abra logs\historico.html no navegador para ver o log visual
echo.
pause
