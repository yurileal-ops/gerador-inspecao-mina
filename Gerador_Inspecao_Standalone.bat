@echo off
REM Caminho absoluto da pasta do gerador
set SCRIPT_DIR=C:\Users\USUARIO\Desktop\gerador_inspecao

REM Verifica se a pasta existe
if not exist "%SCRIPT_DIR%" (
    echo Erro: Pasta do gerador nao encontrada em %SCRIPT_DIR%
    echo Ajuste o caminho neste arquivo .bat
    pause
    exit /b 1
)

REM Executa o app.py usando Python do ambiente virtual
cd /d "%SCRIPT_DIR%"
"%SCRIPT_DIR%\.venv\Scripts\python.exe" "%SCRIPT_DIR%\app.py"

if errorlevel 1 (
    echo.
    echo Erro ao executar o gerador.
    pause
)
