@echo off
setlocal
set "HERE=%~dp0"
set "PY=%HERE%.venv\Scripts\python.exe"
call :ensure_python_exe PY

if not exist "%PY%" (
  echo Creando entorno...
  python -m venv "%HERE%.venv" || exit /b 1
  set "PY=%HERE%.venv\Scripts\python.exe"
  call :ensure_python_exe PY
  if not exist "%PY%" (
    echo No se encontro el interprete de Python esperado en "%PY%".
    echo Revisa la configuracion de la variable PY al inicio de este script.
    exit /b 9009
  )
  "%PY%" -m pip install -U pip
  "%PY%" -m pip install nicegui pywebview pandas==2.3.2 openpyxl==3.1.5 python-dotenv
)

"%PY%" -m rentabilidad.gui.app
set ERR=%ERRORLEVEL%

if not "%ERR%"=="0" (
  echo Ocurrio un error durante la ejecucion. Codigo: %ERR%
  pause
  exit /b %ERR%
)

pause
goto :eof

:ensure_python_exe
setlocal
set "__var=%~1"
call set "__value=%%%__var%%%"
for %%I in ("%__value%") do (
  if /I "%%~xI"=="" (
    set "__value=%%~fI\python.exe"
  ) else (
    set "__value=%%~fI"
  )
)
endlocal & set "%__var%=%__value%"
exit /b 0
