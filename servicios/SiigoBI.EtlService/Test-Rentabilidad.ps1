[CmdletBinding()]
param([string]$Fecha = '', [switch]$Replace, [switch]$CheckOnly, [switch]$CaptureProducts)
$ErrorActionPreference = 'Stop'
$config = Get-Content 'C:\SiigoBI_ETL\appsettings.json' -Raw | ConvertFrom-Json
$target = $config.SqlTargets.RentReport
if (-not $target) { throw 'Primero instala el paquete.' }
$keys = @('SQL_SERVER','SQL_DATABASE','SQL_USER','SQL_PASSWORD','SQL_TRUSTED','SQL_ENCRYPT','SQL_TRUST_CERT','PYTHONUTF8')
$previous = @{}
foreach ($key in $keys) { $previous[$key] = [Environment]::GetEnvironmentVariable($key, 'Process') }
try {
    $env:SQL_SERVER = $target.Server
    $env:SQL_DATABASE = 'SiigoRent'
    $env:SQL_USER = $target.User
    $env:SQL_PASSWORD = $target.Password
    $env:SQL_TRUSTED = if ($target.TrustedConnection) { '1' } else { '0' }
    $env:SQL_ENCRYPT = '1'
    $env:SQL_TRUST_CERT = '0'
    $env:PYTHONUTF8 = '1'
    $arguments = @('C:\Rentabilidad\Rent\servicios\etl_rentabilidad.py')
    if ($Fecha) { $arguments += @('--fecha', $Fecha) }
    if ($CheckOnly) { $arguments += '--check-sql' }
    if ($CaptureProducts) { $arguments += '--capture-products' }
    if ($Replace) { $arguments += '--replace' }
    & 'C:\Rentabilidad\Rent\.venv-sql\Scripts\python.exe' @arguments
    if ($LASTEXITCODE -ne 0) { throw 'La prueba fallo. No actives el trabajo; revisa el mensaje anterior.' }
    Write-Host 'PRUEBA CORRECTA. Consulta el resultado anterior.'
}
finally {
    foreach ($key in $keys) { [Environment]::SetEnvironmentVariable($key, $previous[$key], 'Process') }
}
