@echo off
setlocal
:: ------------------------------------------------
:: Script: Actualizacion de productos para Rentabilidad
:: Las rutas se cargan desde .env o variables de entorno para evitar
:: dependencias rigidas a unidades locales (D:, Z:, etc.).
:: ------------------------------------------------

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
if not defined SIIGO_COMMAND set "SIIGO_COMMAND=ExcelSIIGO"
if not defined SIIGO_REPORTE set "SIIGO_REPORTE=GETINV"
if not defined SIIGO_EMPRESA set "SIIGO_EMPRESA=L"
if not defined SIIGO_USUARIO set "SIIGO_USUARIO=JUAN"
if not defined SIIGO_CLAVE set "SIIGO_CLAVE=0110"
if not defined SIIGO_ESTADO_PARAM set "SIIGO_ESTADO_PARAM=S"
if not defined SIIGO_RANGO_INI set "SIIGO_RANGO_INI=0010001000001"
if not defined SIIGO_RANGO_FIN set "SIIGO_RANGO_FIN=0400027999999"

for /f %%a in ('powershell -NoProfile -Command "(Get-Date).ToString('yyyy')"') do set "ANO=%%a"
for /f %%b in ('powershell -NoProfile -Command "(Get-Date).ToString('MM')"') do set "MES=%%b"
for /f %%c in ('powershell -NoProfile -Command "(Get-Date).ToString('dd')"') do set "DIA=%%c"

if not exist "%SIIGO_DIR%" (
  echo ERROR: No existe la carpeta de SIIGO "%SIIGO_DIR%". Ajusta SIIGO_DIR en .env o en el entorno.
  pause
  exit /b 2
)

if not exist "%PRODUCTOS_DIR%" mkdir "%PRODUCTOS_DIR%"

cd /d "%SIIGO_DIR%"

:: ===================== PRODUCTOS ============================
"%SIIGO_COMMAND%" "%SIIGO_BASE%\" %ANO% %SIIGO_REPORTE% %SIIGO_EMPRESA% %SIIGO_USUARIO% %SIIGO_CLAVE% "%SIIGO_LOG%" %SIIGO_ESTADO_PARAM% %SIIGO_RANGO_INI% %SIIGO_RANGO_FIN% "%PRODUCTOS_DIR%\Productos%MES%%DIA%.xlsx"
IF %ERRORLEVEL% NEQ 0 EXIT /B %ERRORLEVEL%

echo.
echo OK: Productos actualizados en "%PRODUCTOS_DIR%\Productos%MES%%DIA%.xlsx"
endlocal
