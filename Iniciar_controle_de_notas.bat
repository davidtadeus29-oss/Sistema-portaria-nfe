@echo off
title Diagnostico e Execucao do Sistema de Notas
cd /d "%~dp0"
cls

echo ========================================================
echo 1. PASTA ATUAL E ARQUIVOS ENCONTRADOS:
echo ========================================================
echo Pasta: %CD%
dir /b *.py *.txt
echo.

echo ========================================================
echo 2. TESTANDO VERSAO DO PYTHON:
echo ========================================================
python --version
if errorlevel 1 (
    echo [ERRO] Python nao esta configurado no PATH do Windows!
    goto FIM
)
echo.

echo ========================================================
echo 3. TESTANDO TKINTER E MODULOS:
echo ========================================================
python -c "import tkinter; print('-> Interface Grafica (Tkinter): OK')"
python -c "import openpyxl; print('-> Excel (openpyxl): OK')" 2>nul || python -m pip install openpyxl
python -c "import schedule; print('-> Agendador (schedule): OK')" 2>nul || python -m pip install schedule
echo.

echo ========================================================
echo 4. EXECUTANDO O CODIGO:
echo ========================================================

:: Tenta rodar pelo nome que encontrar
for %%F in (*bipagem*.py *notas*.py Controle*.py app*.py *.py) do (
    if exist "%%F" (
        echo Executando arquivo: "%%F" ...
        echo ----------------------------------------------------
        python "%%F"
        goto FIM
    )
)

echo [ERRO] Nenhum arquivo .py foi encontrado nesta pasta!

:FIM
echo.
echo ========================================================
echo Fim da execucao. Veja as mensagens acima.
echo ========================================================
pause