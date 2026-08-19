@echo off
chcp 65001 >nul
title Sistema de Correspondencia

echo ========================================
echo   Sistema de Correspondencia
echo ========================================
echo.

cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado.
    echo Por favor instala Python desde https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" durante la instalacion.
    pause
    exit /b 1
)

echo [1/4] Python detectado correctamente.

if not exist "venv" (
    echo [2/4] Creando entorno virtual...
    python -m venv venv
) else (
    echo [2/4] Entorno virtual ya existe.
)

call venv\Scripts\activate.bat

echo [3/4] Instalando dependencias...
pip install -r webapp\backend\requirements.txt -q

echo [4/4] Iniciando servidor...
echo.
echo ========================================
echo   Servidor iniciado en: http://localhost:8000
echo   Presiona Ctrl+C para detener
echo ========================================
echo.

start http://localhost:8000

python -m uvicorn webapp.backend.main:app --host 127.0.0.1 --port 8000

pause
