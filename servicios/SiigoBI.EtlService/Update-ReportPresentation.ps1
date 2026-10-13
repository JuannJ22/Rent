[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$python = 'C:\Rentabilidad\Rent\.venv-sql\Scripts\python.exe'
if (-not (Test-Path $python)) { throw 'No se encontro el entorno Python SQL.' }
$root = 'C:\Rentabilidad\Rent'
$backup = Join-Path 'C:\Rentabilidad' ('respaldo-presentacion-' + (Get-Date -Format 'yyyyMMdd-HHmmss-ffff'))
New-Item -ItemType Directory -Path $backup | Out-Null
$origin = 'https://raw.githubusercontent.com/JuannJ22/Rent/859219a'
$files = @('rentabilidad/infra/excel_compat.py','hojas/hoja01_loader.py','servicios/etl_rentabilidad.py')
foreach ($relative in $files) {
    $download = Join-Path $backup ('nuevo-' + (Split-Path $relative -Leaf))
    Invoke-WebRequest "$origin/$relative" -OutFile $download
    @'
import ast, sys
from pathlib import Path
ast.parse(Path(sys.argv[1]).read_text(encoding='utf-8-sig'))
'@ | & $python - $download
    if ($LASTEXITCODE -ne 0) { throw "Archivo Python invalido: $relative. No se instalo la actualizacion." }
    $destination = Join-Path $root $relative
    $saved = Join-Path (Join-Path $backup 'codigo') $relative
    New-Item -ItemType Directory -Path (Split-Path $saved -Parent) -Force | Out-Null
    if (Test-Path $destination) { Copy-Item $destination $saved }
}
foreach ($relative in $files) {
    $destination = Join-Path $root $relative
    $download = Join-Path $backup ('nuevo-' + (Split-Path $relative -Leaf))
    Copy-Item $download ($destination + '.compat.tmp')
    Move-Item ($destination + '.compat.tmp') $destination -Force
}
$env:RENT_COMPAT_ROOT = $root
$env:RENT_COMPAT_BACKUP = $backup
try {
@'
import os, shutil, sys
from pathlib import Path
sys.path.insert(0, os.environ['RENT_COMPAT_ROOT'])
from rentabilidad.infra.excel_compat import normalize_excel_application, format_report_presentation
from openpyxl import load_workbook
from uuid import uuid4
root = Path(r'C:\Rentabilidad\InformesDiarios')
backup = Path(os.environ['RENT_COMPAT_BACKUP']) / 'informes'
failures = []
count = 0
for path in sorted(root.rglob('*.xlsx')):
    if any(part.startswith('.rent-') for part in path.parts) or path.name.startswith('~$'):
        continue
    try:
        saved = backup / path.relative_to(root)
        saved.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, saved)
        fresh = path.with_name(path.stem + '.nuevo-' + uuid4().hex + '.xlsx')
        staged = saved.with_suffix('.ajustado.xlsx')
        try:
            wb = load_workbook(path)
            format_report_presentation(wb)
            wb.save(staged)
            wb.close()
            normalize_excel_application(staged)
            shutil.copyfile(staged, fresh)
            os.replace(fresh, path)
        finally:
            fresh.unlink(missing_ok=True)
            staged.unlink(missing_ok=True)
        count += 1
        print('AJUSTADO: ' + str(path))
    except Exception as error:
        failures.append(path)
        print(f'PENDIENTE: {path}: {error}')
print(f'Informes corregidos: {count}. Pendientes: {len(failures)}')
if failures:
    sys.exit(1)
'@ | & $python -
    if ($LASTEXITCODE -ne 0) { throw 'Hay informes pendientes. Cierra esos archivos en Excel y vuelve a ejecutar el script.' }
}
finally {
    Remove-Item Env:RENT_COMPAT_ROOT, Env:RENT_COMPAT_BACKUP -ErrorAction SilentlyContinue
    Write-Host 'Respaldo:' $backup
}
Write-Host 'Generador actualizado. No se reinicio ni se cambio la configuracion del servicio ETL.'
