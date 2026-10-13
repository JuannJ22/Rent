$ErrorActionPreference = 'Stop'
$path = 'C:\SiigoBI_ETL\appsettings.json'
$config = Get-Content $path -Raw | ConvertFrom-Json
$job = @($config.Jobs | Where-Object { $_.Name -in @('RentabilidadDaily','ProductosDaily') })
if ($job.Count -ne 2) { throw 'No hay un unico trabajo RentabilidadDaily instalado.' }
$state = Get-Content $config.Service.StateFile -Raw | ConvertFrom-Json
if (@($state.Jobs.PSObject.Properties | Where-Object { $_.Value.IsRunning }).Count -gt 0) {
    throw 'El ETL esta trabajando. Repite cuando termine; no se modifico nada.'
}
$backup = $path + '.antes-rentabilidad-' + (Get-Date -Format 'yyyyMMdd-HHmmss-ffff')
Copy-Item $path $backup
foreach ($item in $job) { $item.Enabled = $true }
Stop-Service 'SiigoBI.EtlService' -ErrorAction Stop
try {
    [IO.File]::WriteAllText($path, ($config | ConvertTo-Json -Depth 40), [Text.UTF8Encoding]::new($false))
    Start-Service 'SiigoBI.EtlService' -ErrorAction Stop
    Start-Sleep -Seconds 3
    if ((Get-Service 'SiigoBI.EtlService').Status -ne 'Running') { throw 'El servicio no se mantuvo activo; se restaurara la configuracion.' }
}
catch {
    $failure = $_
    Copy-Item $backup $path -Force
    Start-Service 'SiigoBI.EtlService' -ErrorAction SilentlyContinue
    throw $failure
}
Write-Host 'ACTIVADOS: ProductosDaily a las 20:45 y RentabilidadDaily a las 05:00 Colombia.'
Write-Host 'El primer informe corresponde al primer dia con captura nocturna. Archivos: C:\Rentabilidad\Productos.'
Get-Service 'SiigoBI.EtlService' | Format-Table Name, Status
