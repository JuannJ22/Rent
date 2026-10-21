[CmdletBinding()]
param([string]$Fecha = '2026-10-07', [string]$Productos = '')
$ErrorActionPreference = 'Stop'
$python = 'C:\Rentabilidad\Rent\.venv-sql\Scripts\python.exe'
$loader = 'C:\Rentabilidad\Rent\hojas\hoja01_loader.py'
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss-ffff'
$backup = "C:\Rentabilidad\_Archivo\respaldos\cantidades-$stamp"
$review = "C:\Rentabilidad\_Revision\cantidades-$stamp"
$origin = 'https://raw.githubusercontent.com/JuannJ22/Rent/5fb946b'
$state = Get-Content 'C:\SiigoBI_ETL\state.json' -Raw | ConvertFrom-Json
foreach ($name in @('ProductosDaily','RentabilidadDaily')) {
    if ($state.Jobs.$name.IsRunning) { throw "$name esta activo. Espera a que termine y vuelve a ejecutar; no se modifico nada." }
}
New-Item -ItemType Directory -Path $backup -Force | Out-Null
$helper = Join-Path $backup 'fix_repeated_quantities.py'
$download = Join-Path $backup 'hoja01_loader-nuevo.py'
Invoke-WebRequest "$origin/servicios/fix_repeated_quantities.py" -OutFile $helper
Invoke-WebRequest "$origin/hojas/hoja01_loader.py" -OutFile $download
@'
import ast,sys
from pathlib import Path
for arg in sys.argv[1:]: ast.parse(Path(arg).read_text(encoding='utf-8-sig'))
'@ | & $python - $helper $download
if ($LASTEXITCODE -ne 0) { throw 'No se pudieron validar los scripts. No se instalo nada.' }
Copy-Item $loader (Join-Path $backup 'hoja01_loader-original.py')
& $python $helper --apply-sql --backup $backup
if ($LASTEXITCODE -ne 0) { throw 'No se confirmo la migracion SQL. Revisa el resultado anterior.' }
try {
    Copy-Item $download ($loader + '.nuevo')
    Move-Item ($loader + '.nuevo') $loader -Force
}
catch {
    Copy-Item (Join-Path $backup 'hoja01_loader-original.py') $loader -Force
    throw 'Las cantidades SQL quedaron corregidas, pero no se actualizo el aviso de terceros. Revisa el error de copia.'
}
$arguments = @($helper,'--review-dir',$review,'--fecha',$Fecha)
if ($Productos) { $arguments += @('--productos',$Productos) }
& $python @arguments
if ($LASTEXITCODE -ne 0) { throw "Las cantidades SQL quedaron corregidas, pero falta la copia de revision. Respaldo: $backup" }
Write-Host 'CORRECTO: cantidades sumadas solo en ventas; terceros generales sin aviso de vendedor.'
Write-Host 'Respaldo SQL y codigo:' $backup
Write-Host 'Revision:' $review
Write-Host 'Los proximos informes usaran la correccion. Los ya publicados se conservaron.'
