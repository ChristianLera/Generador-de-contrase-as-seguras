# Script de ejecución para PowerShell
# Autor: Christian Lera
# Descripción: Gestor profesional de contraseñas seguras

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "   🔐 GESTOR PROFESIONAL DE CONTRASEÑAS SEGURAS" -ForegroundColor Yellow
Write-Host "                 Christian Lera" -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Verificar Python
Write-Host "Verificando Python..." -ForegroundColor Green
try {
    $pythonVersion = python --version 2>&1
    Write-Host "[OK] $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Python no está instalado en el sistema." -ForegroundColor Red
    Write-Host ""
    Write-Host "Por favor, instala Python 3.8 o superior desde:" -ForegroundColor Yellow
    Write-Host "https://www.python.org/downloads/" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Asegurate de marcar 'Add Python to PATH' durante la instalación." -ForegroundColor Yellow
    Read-Host "Presiona Enter para salir"
    exit 1
}

Write-Host ""
Write-Host "Instalando/verificando dependencias..." -ForegroundColor Green

# Instalar dependencias
$dependencies = @("pandas", "openpyxl", "pyperclip")
$failed = $false

foreach ($dep in $dependencies) {
    Write-Host "  - $dep" -ForegroundColor Gray
    $result = python -m pip install $dep -q 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "    [ERROR] No se pudo instalar $dep" -ForegroundColor Red
        $failed = $true
    }
}

if ($failed) {
    Write-Host ""
    Write-Host "[ERROR] No se pudieron instalar todas las dependencias." -ForegroundColor Red
    Write-Host ""
    Write-Host "Intenta manualmente con:" -ForegroundColor Yellow
    Write-Host "pip install pandas openpyxl pyperclip" -ForegroundColor Cyan
    Read-Host "Presiona Enter para salir"
    exit 1
}

Write-Host ""
Write-Host "[OK] Dependencias listas" -ForegroundColor Green
Write-Host ""
Write-Host "Iniciando la aplicación..." -ForegroundColor Yellow
Write-Host ""

# Ejecutar el programa
python GeneradorDeContraseñasSeguras.py

# Si el programa se cierra
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "   La aplicación se ha cerrado" -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Cyan
Read-Host "Presiona Enter para salir"