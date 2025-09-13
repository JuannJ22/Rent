@echo off
setlocal
set "PROJ_DIR=%~dp0"
if exist "%PROJ_DIR%.env" (
  for /f "usebackq tokens=1* delims==" %%a in ("%PROJ_DIR%.env") do (
    if not "%%a"=="" set "%%a=%%b"
  )
)
if not defined EXCZDIR set "EXCZDIR=D:\SIIWI01\LISTADOS"

where python >nul 2>nul || (echo ERROR: Python no esta en PATH.& pause & exit /b 9001)

if "%~1"=="" (
  python "%PROJ_DIR%hojas\hoja01_loader.py" --exczdir "%EXCZDIR%"
) else (
  python "%PROJ_DIR%hojas\hoja01_loader.py" --excel "%~1" --exczdir "%EXCZDIR%"
)

set ERR=%ERRORLEVEL%
if not "%ERR%"=="0" (echo ERROR: Loader fallo con codigo %ERR% & pause & exit /b %ERR%)
echo OK.
pause
endlocal
