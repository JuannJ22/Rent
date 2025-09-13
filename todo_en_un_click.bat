@echo off
setlocal enabledelayedexpansion
REM ===== CONFIG =====
set "RENT_DIR=C:\Rentabilidad"
set "TEMPLATE=%RENT_DIR%\PLANTILLA.xlsx"
set "EXCZDIR=D:\SIIWI01\LISTADOS"
REM ==================
set "PROJ_DIR=%~dp0"

where python >nul 2>nul || (
  echo ERROR: Python no esta en PATH. Instala Python y reintenta.
  pause
  exit /b 9001
)

if not exist "%RENT_DIR%" mkdir "%RENT_DIR%"

if not exist "%TEMPLATE%" (
  echo ERROR: No existe la plantilla "%TEMPLATE%".
  echo Copia tu PLANTILLA.xlsx a esa ruta y reintenta.
  pause
  exit /b 2
)

if exist "%PROJ_DIR%requirements.txt" (
  echo Instalando/actualizando dependencias...
  python -m pip install -r "%PROJ_DIR%requirements.txt"
)

for /f "usebackq delims=" %%F in (`
  python "%PROJ_DIR%excel_base\clone_from_template.py" --template "%TEMPLATE%" --outdir "%RENT_DIR%"
`) do set "INFORME=%%F"

if not defined INFORME (
  echo ERROR: No se pudo clonar la plantilla.
  pause
  exit /b 3
)
echo INFO: Informe creado: "%INFORME%"

if not exist "%EXCZDIR%" (
  echo ADVERTENCIA: No existe "%EXCZDIR%". Cambia EXCZDIR en este .bat si tu ruta es otra.
  pause
  exit /b 4
)

python "%PROJ_DIR%hojas\hoja01_loader.py" --excel "%INFORME%" --exczdir "%EXCZDIR%"
set ERR=%ERRORLEVEL%
if not "%ERR%"=="0" (
  echo ERROR: Loader fallo con codigo %ERR%
  pause
  exit /b %ERR%
)

echo.
echo ✅ OK: Proceso completado. Archivo final:
echo    "%INFORME%"
echo.
pause
endlocal
