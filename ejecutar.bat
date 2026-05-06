@echo off
title 🔐 Gestor de Contraseñas Seguras - Christian Lera
echo ================================================
echo    🔐 GESTOR PROFESIONAL DE CONTRASEÑAS SEGURAS
echo                  Christian Lera
echo ================================================
echo.
echo Verificando dependencias...
echo.

:: Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado en el sistema.
    echo.
    echo Por favor, instala Python 3.8 o superior desde:
    echo https://www.python.org/downloads/
    echo.
    echo Asegurate de marcar "Add Python to PATH" durante la instalacion.
    pause
    exit /b 1
)

:: Mostrar versión de Python
python --version
echo.

:: Verificar e instalar dependencias
echo Instalando/verificando dependencias...
pip install pandas openpyxl pyperclip -q

if errorlevel 1 (
    echo [ERROR] No se pudieron instalar las dependencias.
    echo.
    echo Intenta manualmente con:
    echo pip install pandas openpyxl pyperclip
    pause
    exit /b 1
)

echo.
echo [OK] Dependencias listas
echo.
echo Iniciando la aplicacion...
echo.

:: Ejecutar el programa
python GeneradorDeContraseñasSeguras.py

:: Si el programa se cierra, mostrar mensaje
echo.
echo ================================================
echo    La aplicacion se ha cerrado
echo ================================================
pause