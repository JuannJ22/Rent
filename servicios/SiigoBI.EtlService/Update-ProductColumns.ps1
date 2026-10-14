[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$python = 'C:\Rentabilidad\Rent\.venv-sql\Scripts\python.exe'
$module = 'C:\Rentabilidad\Rent\rentabilidad\infra\product_snapshots.py'
$backup = 'C:\Rentabilidad\respaldo-columnas-productos-' + (Get-Date -Format 'yyyyMMdd-HHmmss-ffff')
New-Item -ItemType Directory -Path $backup | Out-Null
$download = Join-Path $backup 'product_snapshots-nuevo.py'
Invoke-WebRequest 'https://raw.githubusercontent.com/JuannJ22/Rent/1930a43/rentabilidad/infra/product_snapshots.py' -OutFile $download
@'
import ast,sys
from pathlib import Path
ast.parse(Path(sys.argv[1]).read_text(encoding='utf-8-sig'))
'@ | & $python - $download
if ($LASTEXITCODE -ne 0) { throw 'No se pudo validar la actualizacion. No se instalo.' }
Copy-Item $module (Join-Path $backup 'product_snapshots-original.py')
Copy-Item $download ($module + '.nuevo')
Move-Item ($module + '.nuevo') $module -Force
try {
@'
import sys, shutil
from pathlib import Path
from openpyxl import load_workbook
sys.path.insert(0, r'C:\Rentabilidad\Rent')
from rentabilidad.infra.product_snapshots import compact_snapshot
backup = Path(sys.argv[1]) / 'productos'
failures = []
for path in sorted(Path(r'C:\Rentabilidad\Productos').glob('productos-*.xlsx')):
    try:
        wb = load_workbook(path, read_only=True)
        try:
            owned = 'PRODUCTOS' in wb.sheetnames and 'CAPTURA' in wb.sheetnames
        finally:
            wb.close()
        if not owned:
            continue
        backup.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, backup / path.name)
        if compact_snapshot(path):
            print('PRODUCTOS CORREGIDOS: ' + str(path))
        else:
            print('PRODUCTOS YA CORRECTOS: ' + str(path))
    except Exception as error:
        failures.append(path)
        print(f'PENDIENTE: {path}: {error}')
if failures:
    sys.exit(1)
'@ | & $python - $backup
    if ($LASTEXITCODE -ne 0) { throw 'Hay capturas pendientes. Cierra los archivos de productos en Excel y revisa el resultado.' }
    $report = 'C:\Rentabilidad\InformesDiarios\2026\Octubre\Octubre 13.xlsx'
    $prices = 'C:\Rentabilidad\Productos\productos-2026-10-10.xlsx'
    if ((Test-Path $report) -and (Test-Path $prices)) {
        Copy-Item $report (Join-Path $backup 'Octubre 13-antes.xlsx')
        $oldTimeout = $env:SQL_TIMEOUT
        $oldProducts = $env:SQL_LEGACY_PRODUCT_FILE
        try {
            $env:SQL_TIMEOUT = '300'
            $env:SQL_LEGACY_PRODUCT_FILE = $prices
            & 'C:\Rentabilidad\ActualizarETL\Test-Rentabilidad.ps1' -Fecha '2026-10-13' -Replace
        }
        finally {
            $env:SQL_TIMEOUT = $oldTimeout
            $env:SQL_LEGACY_PRODUCT_FILE = $oldProducts
        }
    }
    else {
        Write-Warning 'No se regenero el informe del 13: falta el informe o su listado del 10.'
    }
    Write-Host 'Actualizacion completada. Productos: solo nombre y doce listas de precios.'
}
finally {
    Write-Host 'Respaldo:' $backup
}
