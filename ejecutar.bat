@echo off
title Consulta tiempos Amazon

echo ===========================================
echo        CONSULTA DE TIEMPOS AMAZON
echo ===========================================
echo.

set /p URL=Pega la URL del producto y pulsa ENTER:

if "%URL%"=="" (
    echo.
    echo No se ha introducido ninguna URL.
    pause
    exit
)

echo.
echo Iniciando consulta...
echo.

python main.py "%URL%"

echo.
echo ===========================================
echo Consulta finalizada.
echo ===========================================
pause