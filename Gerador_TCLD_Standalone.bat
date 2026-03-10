@echo off
REM Gerador de Inspeção Semanal - TCLD
REM Caminho absoluto da pasta do gerador
set SCRIPT_DIR=C:\Users\USUARIO\Desktop\gerador_inspecao

REM Verifica se a pasta existe
if not exist "%SCRIPT_DIR%" (
    echo Erro: Pasta do gerador nao encontrada em %SCRIPT_DIR%
    echo Ajuste o caminho neste arquivo .bat
    pause
    exit /b 1
)

REM Executa o app_tcld.py
cd /d "%SCRIPT_DIR%"
"%SCRIPT_DIR%\.venv\Scripts\python.exe" "%SCRIPT_DIR%\app_tcld.py"

if errorlevel 1 (
    echo.
    echo Erro ao executar o gerador TCLD.
    pause
)
