@echo off
setlocal enabledelayedexpansion

REM ============================================
REM Script BAT para ejecutar descarga de Productos Fertrac
REM Con logging completo y manejo de errores
REM ============================================

REM ============================================
REM Configuración de rutas
REM ============================================

REM Obtener la ruta del directorio donde está este BAT
set "SCRIPT_DIR=%~dp0"

REM Ruta del script Python (en el mismo directorio que el BAT)
set "SCRIPT_PY=%SCRIPT_DIR%descargar_inventario_general.py"

REM Ruta de Python (intentar encontrar automáticamente o especificar)
set "PYTHON_EXE=python.exe"

REM Carpeta de logs
set "LOG_DIR=%SCRIPT_DIR%logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM Archivo de log con fecha y hora
set "FECHA=%date:~-4%%date:~3,2%%date:~0,2%"
set "HORA=%time:~0,2%%time:~3,2%%time:~6,2%"
set "HORA=%HORA: =0%"
set "LOG_BAT=%LOG_DIR%\ejecucion_%FECHA%_%HORA%.log"

REM ============================================
REM Inicio de logging
REM ============================================

echo ============================================= > "%LOG_BAT%"
echo INICIO DE EJECUCION - DESCARGA PRODUCTOS FERTRAC >> "%LOG_BAT%"
echo ============================================= >> "%LOG_BAT%"
echo [BAT %date% %time%] Script BAT iniciado >> "%LOG_BAT%"
echo [BAT %date% %time%] Carpeta de trabajo: %SCRIPT_DIR% >> "%LOG_BAT%"
echo [BAT %date% %time%] Archivo de log: %LOG_BAT% >> "%LOG_BAT%"
echo. >> "%LOG_BAT%"

REM ============================================
REM Verificación de Python
REM ============================================

echo [BAT %date% %time%] Verificando Python... >> "%LOG_BAT%"

REM Intentar ejecutar python
"%PYTHON_EXE%" --version >> "%LOG_BAT%" 2>&1
if !ERRORLEVEL! NEQ 0 (
    echo [BAT %date% %time%] ERROR: Python no encontrado en PATH >> "%LOG_BAT%"
    echo [BAT %date% %time%] Intentando buscar Python en ubicaciones comunes... >> "%LOG_BAT%"
    
    REM Intentar ubicaciones comunes
    if exist "C:\Python39\python.exe" set "PYTHON_EXE=C:\Python39\python.exe"
    if exist "C:\Python310\python.exe" set "PYTHON_EXE=C:\Python310\python.exe"
    if exist "C:\Python311\python.exe" set "PYTHON_EXE=C:\Python311\python.exe"
    if exist "C:\Python312\python.exe" set "PYTHON_EXE=C:\Python312\python.exe"
    if exist "C:\Python313\python.exe" set "PYTHON_EXE=C:\Python313\python.exe"
    if exist "C:\Python314\python.exe" set "PYTHON_EXE=C:\Python314\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python39\python.exe" set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
    
    "%PYTHON_EXE%" --version >> "%LOG_BAT%" 2>&1
    if !ERRORLEVEL! NEQ 0 (
        echo [BAT %date% %time%] ERROR CRITICO: No se pudo encontrar Python >> "%LOG_BAT%"
        echo [BAT %date% %time%] Por favor instale Python o agregue al PATH >> "%LOG_BAT%"
        exit /b 1
    )
)

echo [BAT %date% %time%] OK: Python encontrado en %PYTHON_EXE% >> "%LOG_BAT%"
echo. >> "%LOG_BAT%"

REM ============================================
REM Verificación del script Python
REM ============================================

echo [BAT %date% %time%] Verificando script Python... >> "%LOG_BAT%"
echo [BAT %date% %time%] Ruta esperada: %SCRIPT_PY% >> "%LOG_BAT%"
echo. >> "%LOG_BAT%"

REM Listar archivos .py en el directorio para debug
echo [BAT %date% %time%] Archivos .py en el directorio: >> "%LOG_BAT%"
dir "%SCRIPT_DIR%\*.py" /b >> "%LOG_BAT%" 2>&1
echo. >> "%LOG_BAT%"

if not exist "%SCRIPT_PY%" (
    echo [BAT %date% %time%] ERROR: No existe script Python >> "%LOG_BAT%"
    echo [BAT %date% %time%] Intentando buscar el script con nombre similar... >> "%LOG_BAT%"
    
    REM Buscar archivos que contengan "productos" y "fertrac"
    for %%F in ("%SCRIPT_DIR%\*productos*.py") do (
        echo [BAT %date% %time%] Encontrado: %%F >> "%LOG_BAT%"
        set "SCRIPT_PY=%%F"
        goto :script_found
    )
    
    echo [BAT %date% %time%] ERROR: No se encontro ningun script >> "%LOG_BAT%"
    echo [BAT %date% %time%] Por favor verifique que el archivo existe >> "%LOG_BAT%"
    exit /b 2
)

:script_found
echo [BAT %date% %time%] OK: Script Python encontrado >> "%LOG_BAT%"
echo [BAT %date% %time%] Ruta: %SCRIPT_PY% >> "%LOG_BAT%"
echo. >> "%LOG_BAT%"

REM ============================================
REM Cambiar al directorio del script
REM ============================================

echo [BAT %date% %time%] Cambiando al directorio del script... >> "%LOG_BAT%"
cd /d "%SCRIPT_DIR%"
echo [BAT %date% %time%] Directorio actual: %CD% >> "%LOG_BAT%"
echo. >> "%LOG_BAT%"

REM ============================================
REM Ejecutar script Python
REM ============================================

echo [BAT %date% %time%] ========================================= >> "%LOG_BAT%"
echo [BAT %date% %time%] EJECUTANDO SCRIPT PYTHON >> "%LOG_BAT%"
echo [BAT %date% %time%] ========================================= >> "%LOG_BAT%"
echo. >> "%LOG_BAT%"

"%PYTHON_EXE%" "%SCRIPT_PY%" >> "%LOG_BAT%" 2>&1
set "PYERR=!ERRORLEVEL!"

echo. >> "%LOG_BAT%"
echo [BAT %date% %time%] ========================================= >> "%LOG_BAT%"
echo [BAT %date% %time%] SCRIPT PYTHON TERMINADO >> "%LOG_BAT%"
echo [BAT %date% %time%] ========================================= >> "%LOG_BAT%"
echo [BAT %date% %time%] Codigo de salida: !PYERR! >> "%LOG_BAT%"
echo. >> "%LOG_BAT%"

if !PYERR! EQU 0 (
    echo [BAT %date% %time%] *** EXITO: Script completado exitosamente *** >> "%LOG_BAT%"
) else (
    echo [BAT %date% %time%] *** ERROR: Script fallo con codigo !PYERR! *** >> "%LOG_BAT%"
    echo [BAT %date% %time%] Revise el log para mas detalles >> "%LOG_BAT%"
)

echo. >> "%LOG_BAT%"
echo [BAT %date% %time%] FIN DE EJECUCION >> "%LOG_BAT%"
echo ============================================= >> "%LOG_BAT%"

endlocal
exit /b %PYERR%