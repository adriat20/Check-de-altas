@echo off
REM ======================================================================
REM  Lanzador universal de checks_altasFILTRO_FIRMA_X.py
REM  → No importa dónde esté este .bat: busca el .py por su ruta relativa
REM ======================================================================

:: 1)  Ruta donde vive el script (.py)
set "SCRIPT_DIR=%USERPROFILE%\OneDrive\ESCRITORIO IBERDROLA\PROGRAMACION\Proyecto_Check_Altas"
set "SCRIPT=%SCRIPT_DIR%\checks_altasFILTRO_FIRMA_OFICI.py"

:: 2)  Comprobaciones básicas
if not exist "%SCRIPT_DIR%" (
    echo ERROR: La carpeta %SCRIPT_DIR% no existe ^(¿OneDrive sin sincronizar?^)
    pause
    exit /b 1
)
if not exist "%SCRIPT%" (
    echo ERROR: No se encuentra %SCRIPT%
    dir /b "%SCRIPT_DIR%"
    pause
    exit /b 1
)

:: 3)  Tomar fechas si no vienen como argumentos
if "%~2"=="" (
    echo Introduce la fecha inicial ^(dd-mm-aaaa^):
    set /p D_INI=
    echo Introduce la fecha final   ^(dd-mm-aaaa^):
    set /p D_FIN=
) else (
    set "D_INI=%~1"
    set "D_FIN=%~2"
)

:: 4)  Ejecutar el script
echo.
echo -----------------------------------------------------------
echo Ejecutando %SCRIPT% %D_INI% %D_FIN%
echo -----------------------------------------------------------
python "%SCRIPT%" %D_INI% %D_FIN%

if errorlevel 1 (
    echo.
    echo *** El script devolvió un error. Revisa los mensajes de arriba. ***
    pause
    exit /b 1
)

:: 5)  Abrir la carpeta de salida
explorer "%SCRIPT_DIR%"

echo.
echo Proceso completado. Pulsa una tecla para salir
pause >nul
