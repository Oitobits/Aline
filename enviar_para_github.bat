@echo off
title Enviar NutriFill AI para o GitHub
cls
echo ==================================================
echo      Enviar NutriFill AI para o GitHub
echo ==================================================
echo.
echo IMPORTANTE: Antes de continuar, voce deve criar um repositorio
echo no seu GitHub (recomendamos criar como PRIVADO).
echo.
echo Nome sugerido do repositorio: nutrifill-ai
echo link para criar: https://github.com/new
echo.
echo ==================================================
echo.

:: Solicita o nome de usuario do GitHub
set /p github_user="Digite o seu nome de usuario do GitHub: "
if "%github_user%"=="" (
    echo [ERRO] O nome de usuario nao pode ser vazio.
    pause
    exit /b 1
)

:: Solicita o nome do repositorio
set /p repo_name="Digite o nome do repositorio criado (padrao: nutrifill-ai): "
if "%repo_name%"=="" (
    set repo_name=nutrifill-ai
)

echo.
echo Vinculando repositorio local ao GitHub...
:: Remove remote se ja existir
git remote remove origin >nul 2>&1
git remote add origin https://github.com/%github_user%/%repo_name%.git

echo.
echo Enviando arquivos...
echo (Se for a primeira vez, o Windows abrira uma janela para voce fazer login no GitHub)
echo.

git push -u origin main

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao enviar para o GitHub.
    echo Verifique se o repositorio foi criado no site com o nome exato
    echo e tente novamente.
    echo.
) else (
    echo.
    echo ==================================================
    echo    Sucesso! Codigo enviado para o GitHub!
    echo ==================================================
    echo Link do seu repositorio: https://github.com/%github_user%/%repo_name%
    echo.
)

pause
