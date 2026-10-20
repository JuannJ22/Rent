[CmdletBinding()]
param([string]$Fecha = ([DateTime]::UtcNow.AddHours(-5).AddDays(-1).ToString('yyyy-MM-dd')))
$ErrorActionPreference = 'Stop'
$python = 'C:\Rentabilidad\Rent\.venv-sql\Scripts\python.exe'
$root = 'C:\Rentabilidad\Rent'
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss-ffff'
$backup = "C:\Rentabilidad\_Archivo\respaldos\validacion-$stamp"
$review = "C:\Rentabilidad\_Revision\sucursales-$stamp"
$origin = 'https://raw.githubusercontent.com/JuannJ22/Rent/3643086'
$files = @('hojas/hoja01_loader.py', 'servicios/etl_rentabilidad.py', 'docs/sql/rentabilidad_cliente_sucursales.sql')
$state = Get-Content 'C:\SiigoBI_ETL\state.json' -Raw | ConvertFrom-Json
foreach ($name in @('ProductosDaily', 'RentabilidadDaily')) {
    if ($state.Jobs.$name.IsRunning) { throw "$name esta activo. Espera a que termine y vuelve a ejecutar; no se modifico nada." }
}
New-Item -ItemType Directory -Path $backup -Force | Out-Null
foreach ($relative in $files) {
    $download = Join-Path $backup ('nuevo-' + (Split-Path $relative -Leaf))
    Invoke-WebRequest "$origin/$relative" -OutFile $download
    if ($relative.EndsWith('.py')) {
@'
import ast, sys
from pathlib import Path
ast.parse(Path(sys.argv[1]).read_text(encoding='utf-8-sig'))
'@ | & $python - $download
        if ($LASTEXITCODE -ne 0) { throw "No se pudo validar $relative. No se instalo nada." }
        $saved = Join-Path (Join-Path $backup 'codigo') $relative
        New-Item -ItemType Directory -Path (Split-Path $saved -Parent) -Force | Out-Null
        Copy-Item (Join-Path $root $relative) $saved
    }
}
# La vista nueva se compila antes de activar el loader. No se imprimen credenciales.
@'
import json, sys
from pathlib import Path
sys.path.insert(0, r'C:\Rentabilidad\Rent')
import pyodbc
from rentabilidad.infra.sql_server import SqlServerConfig
backup = Path(sys.argv[1])
def option(obj, name, default=None):
    return next((value for key, value in obj.items() if key.lower() == name.lower()), default)
settings = json.loads(Path(r'C:\SiigoBI_ETL\appsettings.json').read_text(encoding='utf-8-sig'))
target = option(option(settings, 'SqlTargets', {}), 'RentReport')
if not target: raise RuntimeError('Falta SqlTargets.RentReport.')
config = SqlServerConfig(server=option(target,'Server'), database='SiigoRent',
    user=option(target,'User'), password=option(target,'Password'),
    trusted_connection=bool(option(target,'TrustedConnection',False)),
    driver='ODBC Driver 18 for SQL Server', encrypt=True, trust_server_certificate=False, timeout=300)
conn = pyodbc.connect(config.connection_string(), timeout=30)
conn.timeout = 300
try:
    cursor = conn.cursor()
    try:
        old = cursor.execute("SELECT OBJECT_DEFINITION(OBJECT_ID('dbo.vw_rentabilidad_cliente_informe'))").fetchone()[0]
        if old:
            (backup / 'vista-informe-anterior.sql').write_text(old, encoding='utf-8')
        cursor.execute((backup / 'nuevo-rentabilidad_cliente_sucursales.sql').read_text(encoding='utf-8-sig'))
        cursor.execute('SELECT TOP (0) SUCURSAL, VENTAS_CON_IVA, DOCUMENTO FROM dbo.vw_rentabilidad_cliente_informe')
        cursor.fetchall()
        conn.commit()
        print('SQL: vista del informe creada y columnas comprobadas.')
    finally:
        cursor.close()
except BaseException:
    conn.rollback()
    raise
finally:
    conn.close()
'@ | & $python - $backup
if ($LASTEXITCODE -ne 0) { throw 'La vista no se pudo preparar. Los scripts del servicio siguen intactos.' }
try {
    foreach ($relative in $files | Where-Object { $_.EndsWith('.py') }) {
        $destination = Join-Path $root $relative
        $download = Join-Path $backup ('nuevo-' + (Split-Path $relative -Leaf))
        Copy-Item $download ($destination + '.nuevo')
        Move-Item ($destination + '.nuevo') $destination -Force
    }
}
catch {
    foreach ($relative in $files | Where-Object { $_.EndsWith('.py') }) {
        Copy-Item (Join-Path (Join-Path $backup 'codigo') $relative) (Join-Path $root $relative) -Force
    }
    throw
}
Write-Host 'Codigo actualizado. Preparando copia de revision; los informes publicados se conservan.'
@'
import csv, json, os, shutil, subprocess, sys
from pathlib import Path
from datetime import date
root = Path(r'C:\Rentabilidad\Rent')
sys.path.insert(0, str(root))
review = Path(sys.argv[1])
day = date.fromisoformat(sys.argv[2])
from rentabilidad.infra.product_snapshots import read_snapshot, snapshot_path
source = snapshot_path(Path(r'C:\Rentabilidad\Productos'), day)
if not source.exists():
    print(f'REVISION PENDIENTE: falta captura de {day}. El codigo ya quedo actualizado. Repite con -Fecha YYYY-MM-DD que tenga captura.')
    sys.exit(0)
read_snapshot(source, day)
copied = snapshot_path(review / 'Productos', day)
copied.parent.mkdir(parents=True, exist_ok=True)
shutil.copyfile(source, copied)
settings = json.loads(Path(r'C:\SiigoBI_ETL\appsettings.json').read_text(encoding='utf-8-sig'))
def option(obj, name, default=None):
    return next((value for key, value in obj.items() if key.lower() == name.lower()), default)
target = option(option(settings,'SqlTargets',{}),'RentReport')
os.environ.update(SQL_SERVER=option(target,'Server'), SQL_DATABASE='SiigoRent',
    SQL_USER=option(target,'User','') or '', SQL_PASSWORD=option(target,'Password','') or '',
    SQL_TRUSTED='1' if option(target,'TrustedConnection',False) else '0',
    SQL_ENCRYPT='1', SQL_TRUST_CERT='0', SQL_TIMEOUT='300', RENT_DIR=str(review),
    PYTHONUTF8='1', PYTHONIOENCODING='utf-8')
os.environ.pop('SQL_LEGACY_PRODUCT_FILE', None)
os.environ.pop('SQL_PRODUCT_SNAPSHOT', None)
result = subprocess.run([sys.executable, str(root / 'servicios' / 'etl_rentabilidad.py'),
    '--fecha', day.isoformat(), '--base-dir', str(review), '--template', r'C:\Rentabilidad\PLANTILLA.xlsx'], cwd=root)
if result.returncode:
    sys.exit(result.returncode)
# Audita negativos y ajustes sin cantidad con el mismo catalogo historico usado por Excel.
from rentabilidad.infra.sql_server import SqlServerConfig, fetch_dataframe
from hojas.hoja01_loader import _normalize_sql_text, _normalize_product_key
from openpyxl import load_workbook
book_path = next((review / 'InformesDiarios').rglob('*.xlsx'))
book = load_workbook(book_path, read_only=True, data_only=True)
try:
    prices = {}
    for row in book['PRECIOS'].iter_rows(max_col=13, values_only=True):
        key = _normalize_product_key(row[0])
        if key and key not in prices: prices[key] = row[12]
finally:
    book.close()
config = SqlServerConfig(server=os.environ['SQL_SERVER'], database='SiigoRent',
    user=os.environ['SQL_USER'], password=os.environ['SQL_PASSWORD'],
    trusted_connection=os.environ['SQL_TRUSTED']=='1', driver='ODBC Driver 18 for SQL Server', timeout=300)
data = _normalize_sql_text(fetch_dataframe(config,
    'SELECT NIT,SUCURSAL,DOCUMENTO,DESCRIPCION,CANTIDAD,VENTAS AS VENTAS,VENTAS_CON_IVA,FILAS_VENTA,CANTIDAD_VENTA_SUM FROM dbo.vw_rentabilidad_cliente_informe WHERE FECHA=?', [day.isoformat()]))
audit = review / 'revision-descuentos.csv'
count = 0
with audit.open('w', encoding='utf-8-sig', newline='') as handle:
    writer = csv.writer(handle, delimiter=';')
    writer.writerow(['Documento','NIT','Sucursal','Producto','Cantidad','Ventas','Ventas con IVA','Lista 12','Descuento %','Filas de ventas','Suma cantidades en ventas','Revision'])
    for row in data.itertuples(index=False):
        price = prices.get(_normalize_product_key(row.DESCRIPCION))
        quantity = float(row.CANTIDAD or 0)
        gross = float(row.VENTAS_CON_IVA or 0)
        discount = None if not quantity or price is None or float(price)<=0 else 1-abs(gross)/abs(quantity)/float(price)
        reason = ('Sin cantidad' if not quantity else 'Sin precio de lista 12' if price is None or float(price)<=0 else 'Unitario superior a lista 12' if discount is not None and discount<0 else '')
        if reason and row.FILAS_VENTA > 1 and abs(float(row.CANTIDAD_VENTA_SUM or 0)) != abs(quantity):
            reason += '; varias filas de venta: revisar cantidad agregada'
        if reason:
            writer.writerow([row.DOCUMENTO,row.NIT,row.SUCURSAL,row.DESCRIPCION,quantity,row.VENTAS,gross,price,
                None if discount is None else round(discount*100,4), row.FILAS_VENTA, row.CANTIDAD_VENTA_SUM, reason])
            count += 1
print(f'REVISION PREPARADA: {book_path}')
print(f'AUDITORIA DE DESCUENTOS: {audit} ({count} casos para revisar)')
'@ | & $python - $review $Fecha
if ($LASTEXITCODE -ne 0) { throw "Codigo actualizado, pero la copia de revision no se completo. Respaldo: $backup" }
Write-Host 'Respaldo:' $backup
Write-Host 'No se reinicio el servicio. Las proximas ejecuciones usaran la correccion.'
