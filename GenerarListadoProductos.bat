@echo off
setlocal
set "PROJ_DIR=%~dp0"

if exist "%PROJ_DIR%.env" (
  for /f "usebackq tokens=1* delims==" %%a in ("%PROJ_DIR%.env") do (
    if not "%%a"=="" set "%%a=%%b"
  )
)

if not defined SIIGO_DIR set "SIIGO_DIR=C:\Siigo"
if not defined SIIGO_BASE set "SIIGO_BASE=D:\SIIWI01"
if not defined PRODUCTOS_DIR set "PRODUCTOS_DIR=C:\Rentabilidad\Productos"
if not defined SIIGO_LOG set "SIIGO_LOG=%SIIGO_BASE%\LOGS\log_catalogos.txt"

where python >nul 2>nul || (
  echo ERROR: Python no esta en PATH.
  pause
  exit /b 9001
)

python "%PROJ_DIR%servicios\generar_listado_productos.py" ^
  --siigo-dir "%SIIGO_DIR%" ^
  --siigo-base "%SIIGO_BASE%" ^
  --productos-dir "%PRODUCTOS_DIR%" ^
  --log "%SIIGO_LOG%"

set ERR=%ERRORLEVEL%
if not "%ERR%"=="0" (
  echo ERROR: GenerarListadoProductos fallo con codigo %ERR%
  pause
  exit /b %ERR%
)

echo.
echo ✅ OK: Listado de productos generado.
echo.
pause
endlocal
