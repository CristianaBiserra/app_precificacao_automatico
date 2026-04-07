@echo off
setlocal
cd /d "%~dp0"
title Assistente de Precificacao KGMLan

echo ==========================================
echo Iniciando Assistente de Precificacao...
echo Pasta: %~dp0
echo ==========================================
echo.

if not exist "%~dp0app_precificacao.py" (
    echo ERRO: arquivo app_precificacao.py nao encontrado.
    goto :fim
)

if not exist "%~dp01 NOVO SIMULADOR_PRECIFICACAO_V2.xlsx" (
    echo ERRO: planilha base nao encontrada.
    echo Arquivo esperado: 1 NOVO SIMULADOR_PRECIFICACAO_V2.xlsx
    goto :fim
)

where python >nul 2>nul
if %errorlevel%==0 (
    echo Executando com python...
    python "%~dp0app_precificacao.py"
    echo.
    echo Codigo de retorno: %errorlevel%
    goto :fim
)

where py >nul 2>nul
if %errorlevel%==0 (
    echo Executando com py...
    py "%~dp0app_precificacao.py"
    echo.
    echo Codigo de retorno: %errorlevel%
    goto :fim
)

echo ERRO: Python nao encontrado no PATH.

:fim
echo.
pause