[CmdletBinding()]
param([string]$EtlDir = 'C:\SiigoBI_ETL', [string]$RentDir = 'C:\Rentabilidad\Rent')
$ErrorActionPreference = 'Stop'
$serviceName = 'SiigoBI.EtlService'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal]$identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Ejecuta PowerShell como administrador.'
}
$manifest = Get-Content (Join-Path $PSScriptRoot 'sha256.json') -Raw | ConvertFrom-Json
foreach ($property in $manifest.PSObject.Properties) {
    $actual = (Get-FileHash (Join-Path $PSScriptRoot $property.Name) -Algorithm SHA256).Hash
    if ($actual -ne $property.Value) { throw "Archivo alterado o incompleto: $($property.Name)" }
}
$configPath = Join-Path $EtlDir 'appsettings.json'
$config = Get-Content $configPath -Raw | ConvertFrom-Json
if (-not $config.SqlTargets.Sql2025Mov) { throw 'No existe SqlTargets.Sql2025Mov en la configuracion instalada.' }
foreach ($name in @('Mov45Daily', 'CatalogosMedios', 'CatalogosPesados')) {
    if (-not ($config.Jobs | Where-Object { $_.Name -eq $name -and $_.Enabled })) {
        throw "Falta una carga previa activa: $name"
    }
}
$statePath = [string]$config.Service.StateFile
if (Test-Path $statePath) {
    $state = Get-Content $statePath -Raw | ConvertFrom-Json
    $busy = @($state.Jobs.PSObject.Properties | Where-Object { $_.Value.IsRunning })
    if ($busy.Count -gt 0) { throw 'El ETL esta trabajando. Repite la instalacion cuando termine; no se modifico nada.' }
}
$backup = Join-Path $EtlDir ('backups\rentabilidad-' + (Get-Date -Format 'yyyyMMdd-HHmmss-ffff'))
New-Item $backup -ItemType Directory -Force | Out-Null
Copy-Item $configPath (Join-Path $backup 'appsettings.json')
$files = @(Get-ChildItem (Join-Path $PSScriptRoot 'publish') -File)
foreach ($file in $files) {
    $existing = Join-Path $EtlDir $file.Name
    if (Test-Path $existing) { Copy-Item $existing (Join-Path $backup $file.Name) }
}
$rentFiles = @('hojas\hoja01_loader.py', 'rentabilidad\infra\sql_server.py', 'servicios\etl_rentabilidad.py', 'rentabilidad\infra\product_snapshots.py')
foreach ($relative in $rentFiles) {
    $existing = Join-Path $RentDir $relative
    if (Test-Path $existing) {
        $saved = Join-Path (Join-Path $backup 'rent') $relative
        New-Item (Split-Path $saved) -ItemType Directory -Force | Out-Null
        Copy-Item $existing $saved
    }
}
if (-not $config.SqlTargets.RentReport) {
    $target = $config.SqlTargets.Sql2025Mov | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $target.Server = '192.168.5.10,14331'
    $target.Database = 'SiigoRent'
    $config.SqlTargets | Add-Member -NotePropertyName 'RentReport' -NotePropertyValue $target
}
if (-not ($config.Jobs | Where-Object Name -eq 'RentabilidadDaily')) {
    $job = Get-Content (Join-Path $PSScriptRoot 'rentabilidad-job.example.json') -Raw | ConvertFrom-Json
    $job.Rentabilidad | Add-Member -NotePropertyName 'FirstReportDate' -NotePropertyValue ((Get-Date).ToUniversalTime().AddHours(-5).ToString('yyyy-MM-dd'))
    $config.Jobs = @($config.Jobs) + @($job)
}
if (-not ($config.Jobs | Where-Object Name -eq 'ProductosDaily')) {
    $productsJob = Get-Content (Join-Path $PSScriptRoot 'productos-job.example.json') -Raw | ConvertFrom-Json
    $config.Jobs = @($config.Jobs) + @($productsJob)
}
Stop-Service $serviceName -ErrorAction Stop
try {
    foreach ($file in $files) { Copy-Item $file.FullName (Join-Path $EtlDir $file.Name) -Force }
    foreach ($relative in $rentFiles) {
        $destination = Join-Path $RentDir $relative
        New-Item (Split-Path $destination) -ItemType Directory -Force | Out-Null
        Copy-Item (Join-Path (Join-Path $PSScriptRoot 'rent') $relative) $destination -Force
    }
    $json = $config | ConvertTo-Json -Depth 40
    [IO.File]::WriteAllText($configPath, $json, [Text.UTF8Encoding]::new($false))
    Start-Service $serviceName -ErrorAction Stop
    Start-Sleep -Seconds 3
    if ((Get-Service $serviceName).Status -ne 'Running') { throw 'El servicio no se mantuvo activo; se restaurara el respaldo.' }
}
catch {
    $failure = $_
    Stop-Service $serviceName -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        $saved = Join-Path $backup $file.Name
        if (Test-Path $saved) { Copy-Item $saved (Join-Path $EtlDir $file.Name) -Force }
    }
    Copy-Item (Join-Path $backup 'appsettings.json') $configPath -Force
    foreach ($relative in $rentFiles) {
        $saved = Join-Path (Join-Path $backup 'rent') $relative
        if (Test-Path $saved) { Copy-Item $saved (Join-Path $RentDir $relative) -Force }
        elseif ($relative -eq 'servicios\etl_rentabilidad.py', 'rentabilidad\infra\product_snapshots.py') { Remove-Item (Join-Path $RentDir $relative) -ErrorAction SilentlyContinue }
    }
    Start-Service $serviceName -ErrorAction SilentlyContinue
    throw $failure
}
Write-Host 'INSTALADO. El trabajo nuevo queda deshabilitado hasta probarlo y activarlo.'
Write-Host 'Respaldo:' $backup
Get-Service $serviceName | Format-Table Name, Status
