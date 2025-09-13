@echo off
setlocal
set "RENT_DIR=C:\Rentabilidad"
set "TEMPLATE=%RENT_DIR%\PLANTILLA.xlsx"
set "PROJ_DIR=%~dp0"

where python >nul 2>nul || (echo ERROR: Python no esta en PATH.& pause & exit /b 9001)
if not exist "%RENT_DIR%" mkdir "%RENT_DIR%"

if not exist "%TEMPLATE%" (
  echo ERROR: No existe "%TEMPLATE%". Copia tu PLANTILLA.xlsx ahi.
  pause
  exit /b 2
)

for /f "usebackq delims=" %%F in (`
  python "%PROJ_DIR%excel_base\clone_from_template.py" --template "%TEMPLATE%" --outdir "%RENT_DIR%"
`) do set "INFORME=%%F"

if not defined INFORME (echo ERROR: No se pudo clonar.& pause & exit /b 3)
echo "%INFORME%"
pause
endlocal
