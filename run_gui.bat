@echo off
setlocal
set "HERE=%~dp0"
set "PY=%HERE%.venv\Scripts\python.exe"

if not exist "%PY%" (
  echo Creando entorno...
  python -m venv "%HERE%.venv" || exit /b 1
  "%PY%" -m pip install -U pip
  "%PY%" -m pip install nicegui pywebview pandas==2.3.2 openpyxl==3.1.5 python-dotenv
)

"%PY%" -m rentabilidad.gui.app
