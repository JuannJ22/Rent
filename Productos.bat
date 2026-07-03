@echo off
:: ------------------------------------------------
:: Script: Actualización de productos para la Rentabilidad
:: Año: 2025
:: Carpeta destino: D:\Rentabilidad\Productos
:: Log: log_catalogos.txt
:: Autor: Juan José Ortiz
:: ------------------------------------------------

:: Obtener el año actual
for /f %%a in ('powershell -command "(Get-Date).ToString('yyyy')"') do set ANO=%%a
for /f %%b in ('powershell -command "(Get-Date).ToString('MM')"') do set MES=%%b
for /f %%c in ('powershell -command "(Get-Date).ToString('dd')"') do set DIA=%%c

cd /d C:\Siigo

:: ===================== PRODUCTOS ============================
ExcelSIIGO D:\SIIWI01\ %ANO% GETINV L JUAN 0110 D:\SIIWI01\LOGS\log_catalogos.txt S 0010001000001 0400027999999 C:\Rentabilidad\Productos\Productos%MES%%DIA%.xlsx
IF %ERRORLEVEL% NEQ 0 EXIT /B %ERRORLEVEL%

echo