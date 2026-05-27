@echo off
title NutriFill AI - Servidor Local
echo ==================================================
echo    Iniciando NutriFill AI - Servidor Local
echo ==================================================

:: Verifica se o Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao foi encontrado no sistema!
    echo Por favor, instale o Python 3.
    pause
    exit /b 1
)

:: Instala as dependencias
echo Verificando e instalando dependencias necessarias...
python -m pip install --upgrade pip
python -m pip install flask pypdf openpyxl google-genai pydantic

:: Abre o navegador
echo Abrindo o navegador em http://127.0.0.1:5000 ...
timeout /t 2 >nul
start http://127.0.0.1:5000

:: Executa a aplicacao
python app.py
pause
