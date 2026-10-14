[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$source = 'C:\Rentabilidad\RentActualizacion\servicios\SiigoBI.EtlService'
if (-not (Test-Path (Join-Path $source 'SiigoBI.EtlService.csproj'))) { throw 'No se encontro el proyecto fuente en RentActualizacion.' }
$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
$dotnetExe = if ($dotnet) { $dotnet.Source } else { 'C:\Rentabilidad\Herramientas\dotnet\dotnet.exe' }
if (-not (Test-Path $dotnetExe)) { throw 'No se encontro el SDK .NET.' }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss-ffff'
$work = "C:\Rentabilidad\ActualizarETL\scheduler-$stamp"
$build = Join-Path $work 'source'
$publish = Join-Path $work 'publish'
New-Item -ItemType Directory -Path $build -Force | Out-Null
Get-ChildItem $source -File | Where-Object Extension -in @('.cs','.csproj') | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $build $_.Name)
}
Invoke-WebRequest 'https://raw.githubusercontent.com/JuannJ22/Rent/9bf1190/servicios/SiigoBI.EtlService/Worker.cs' -OutFile (Join-Path $build 'Worker.cs')
& $dotnetExe publish (Join-Path $build 'SiigoBI.EtlService.csproj') -c Release -r win-x64 --self-contained false -o $publish
if ($LASTEXITCODE -ne 0) { throw 'No se pudo compilar. Servicio instalado intacto.' }
$state = Get-Content 'C:\SiigoBI_ETL\state.json' -Raw | ConvertFrom-Json
$running = @($state.Jobs.PSObject.Properties | Where-Object { $_.Value.IsRunning })
if ($running.Count) { throw ('Hay trabajos activos: ' + ($running.Name -join ', ') + '. Espera a que terminen y vuelve a ejecutar. Servicio intacto.') }
$backup = "C:\SiigoBI_ETL\backups\scheduler-$stamp"
New-Item -ItemType Directory -Path $backup | Out-Null
$files = @('SiigoBI.EtlService.dll','SiigoBI.EtlService.pdb')
foreach ($file in $files) {
    if (-not (Test-Path (Join-Path $publish $file))) { throw "Falta $file en la compilacion. Servicio intacto." }
    if (-not (Test-Path (Join-Path 'C:\SiigoBI_ETL' $file))) { throw "Falta $file instalado. Servicio intacto." }
    Copy-Item (Join-Path 'C:\SiigoBI_ETL' $file) (Join-Path $backup $file)
}
Stop-Service 'SiigoBI.EtlService'
try {
    foreach ($file in $files) { Copy-Item (Join-Path $publish $file) (Join-Path 'C:\SiigoBI_ETL' $file) -Force }
    Start-Service 'SiigoBI.EtlService'
    Start-Sleep -Seconds 3
    if ((Get-Service 'SiigoBI.EtlService').Status -ne 'Running') { throw 'El servicio no permanecio activo.' }
}
catch {
    $failure = $_
    Stop-Service 'SiigoBI.EtlService' -ErrorAction SilentlyContinue
    foreach ($file in $files) { Copy-Item (Join-Path $backup $file) (Join-Path 'C:\SiigoBI_ETL' $file) -Force }
    Start-Service 'SiigoBI.EtlService'
    throw $failure
}
Write-Host 'Planificador actualizado. Configuracion, credenciales, estado y Python conservados.'
Write-Host 'Respaldo:' $backup
Get-Service 'SiigoBI.EtlService' | Format-Table Name, Status
