@echo off
setlocal enabledelayedexpansion
REM ============================================
REM Configuración
REM ============================================
set "CARPETA_PROYECTO=C:\MACROS_COMPRAS\MACRO_VENTAS\Actualizar Inv general"
set "PYTHON_EXE=C:\MACROS_COMPRAS\MACRO_VENTAS\python.exe"
set "SCRIPT_PY=%CARPETA_PROYECTO%\actualizar_inventario_general.py"
set "LOG_BAT=%CARPETA_PROYECTO%\log_actualizar_inventario_general.txt"

REM ============================================
REM Ir a carpeta del proyecto
REM ============================================
cd /d "%CARPETA_PROYECTO%"

REM ============================================
REM Log de inicio
REM ============================================
echo. >> "%LOG_BAT%"
echo ============================================= >> "%LOG_BAT%"
echo [BAT %date% %time%] INICIO EJECUCION >> "%LOG_BAT%"
echo ============================================= >> "%LOG_BAT%"
echo [BAT %date% %time%] Usuario: %USERNAME% >> "%LOG_BAT%"
echo [BAT %date% %time%] Dominio: %USERDOMAIN% >> "%LOG_BAT%"
echo [BAT %date% %time%] Carpeta: %CD% >> "%LOG_BAT%"
echo [BAT %date% %time%] SessionName: %SESSIONNAME% >> "%LOG_BAT%"

REM ============================================
REM Verificar sesión
REM ============================================
echo [BAT %date% %time%] Verificando sesion... >> "%LOG_BAT%"
query session %USERNAME% >> "%LOG_BAT%" 2>&1

REM ============================================
REM Cerrar Excel
REM ============================================
echo [BAT %date% %time%] Cerrando EXCEL.EXE... >> "%LOG_BAT%"
taskkill /F /IM EXCEL.EXE >> "%LOG_BAT%" 2>&1
if errorlevel 1 (
    echo [BAT %date% %time%] EXCEL no estaba abierto >> "%LOG_BAT%"
) else (
    echo [BAT %date% %time%] EXCEL cerrado >> "%LOG_BAT%"
    timeout /t 3 /nobreak > nul
)

REM ============================================
REM Verificar archivos necesarios
REM ============================================
echo [BAT %date% %time%] Verificando archivos... >> "%LOG_BAT%"
if not exist "%PYTHON_EXE%" (
    echo [BAT %date% %time%] ERROR: No existe python.exe en %PYTHON_EXE% >> "%LOG_BAT%"
    exit /b 2
)
echo [BAT %date% %time%] OK: python.exe encontrado >> "%LOG_BAT%"

if not exist "%SCRIPT_PY%" (
    echo [BAT %date% %time%] ERROR: No existe script Python en %SCRIPT_PY% >> "%LOG_BAT%"
    exit /b 3
)
echo [BAT %date% %time%] OK: Script Python encontrado >> "%LOG_BAT%"

REM ============================================
REM Ejecutar script Python
REM ============================================
echo [BAT %date% %time%] ========================================= >> "%LOG_BAT%"
echo [BAT %date% %time%] Ejecutando script Python... >> "%LOG_BAT%"
echo [BAT %date% %time%] ========================================= >> "%LOG_BAT%"

"%PYTHON_EXE%" "%SCRIPT_PY%" >> "%LOG_BAT%" 2>&1

set "PYERR=!ERRORLEVEL!"

echo. >> "%LOG_BAT%"
echo [BAT %date% %time%] ========================================= >> "%LOG_BAT%"
echo [BAT %date% %time%] Python termino con codigo: !PYERR! >> "%LOG_BAT%"
if !PYERR! EQU 0 (
    echo [BAT %date% %time%] EXITO: Script completado >> "%LOG_BAT%"
) else (
    echo [BAT %date% %time%] ERROR: Script fallo >> "%LOG_BAT%"
)
echo [BAT %date% %time%] FIN EJECUCION >> "%LOG_BAT%"
echo ============================================= >> "%LOG_BAT%"

endlocal
exit /b %PYERR%