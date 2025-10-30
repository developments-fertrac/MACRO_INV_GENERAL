@echo off
chcp 65001 >nul
title Comparar Inventarios - Lanzador
setlocal

REM Ruta del script (mismo folder que este .bat)
set "SCRIPT=%~dp0comparar_inv_general.py"

if not exist "%SCRIPT%" (
    echo No se encontro el script: "%SCRIPT%"
    echo Pon este .bat en la misma carpeta que comparar_inv_general.py
    pause
    exit /b 1
)

REM Si hay venv local, usarlo primero
set "VENV_PY=%~dp0venv\Scripts\python.exe"
if exist "%VENV_PY%" (
    echo Usando entorno virtual: "%VENV_PY%"
    "%VENV_PY%" "%SCRIPT%"
    goto :postrun
)

REM Probar con el Python Launcher (py)
where py >nul 2>&1
if %errorlevel%==0 (
    echo Usando py launcher...
    py -3 "%SCRIPT%"
    goto :postrun
)

REM Probar con python en PATH
where python >nul 2>&1
if %errorlevel%==0 (
    echo Usando python del sistema...
    python "%SCRIPT%"
    goto :postrun
)

echo No se encontro Python en el sistema.
echo Instala Python desde https://www.python.org/downloads/ o usa un entorno virtual "venv".
pause
exit /b 1

:postrun
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
    echo Proceso completado correctamente.
) else (
    echo El script termino con codigo de salida %RC%.
)
echo.
pause
endlocal
