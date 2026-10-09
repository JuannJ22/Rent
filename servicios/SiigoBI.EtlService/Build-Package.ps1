[CmdletBinding()]
param([string]$Output = 'C:\Rentabilidad\ActualizarETL')
$ErrorActionPreference = 'Stop'
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
$hasSdk = $false
if ($dotnet) { $hasSdk = @(& $dotnet.Source --list-sdks).Count -gt 0 }
if (-not $hasSdk) {
    $tools = 'C:\Rentabilidad\Herramientas\dotnet'
    New-Item $tools -ItemType Directory -Force | Out-Null
    $installer = Join-Path $tools 'dotnet-install.ps1'
    Invoke-WebRequest 'https://dot.net/v1/dotnet-install.ps1' -OutFile $installer -UseBasicParsing
    & $installer -Version '8.0.419' -InstallDir $tools -NoPath
    if ($LASTEXITCODE -ne 0) { throw 'No se pudo instalar el SDK .NET 8.' }
    $dotnetPath = Join-Path $tools 'dotnet.exe'
} else { $dotnetPath = $dotnet.Source }
New-Item $Output -ItemType Directory -Force | Out-Null
& $dotnetPath publish (Join-Path $PSScriptRoot 'SiigoBI.EtlService.csproj') -c Release -r win-x64 --self-contained false -o (Join-Path $Output 'publish')
if ($LASTEXITCODE -ne 0) { throw 'No se pudo compilar el servicio.' }
foreach ($name in @('Install-Rentabilidad.ps1','Test-Rentabilidad.ps1','Enable-Rentabilidad.ps1','rentabilidad-job.example.json','productos-job.example.json')) {
    Copy-Item (Join-Path $PSScriptRoot $name) (Join-Path $Output $name) -Force
}
foreach ($relative in @('hojas\hoja01_loader.py','rentabilidad\infra\sql_server.py','rentabilidad\infra\product_snapshots.py','servicios\etl_rentabilidad.py')) {
    $target = Join-Path (Join-Path $Output 'rent') $relative
    New-Item (Split-Path $target) -ItemType Directory -Force | Out-Null
    Copy-Item (Join-Path $repo $relative) $target -Force
}
Copy-Item (Join-Path $repo 'docs\sql\zona7_con_total.sql') $Output -Force
Copy-Item (Join-Path $repo 'docs\integracion_siigobi_etl.md') (Join-Path $Output 'LEEME.md') -Force
$manifest = @{}
$base = (Get-Item $Output).FullName.TrimEnd('\')
Get-ChildItem $Output -Recurse -File | Where-Object Name -ne 'sha256.json' | ForEach-Object {
    $relative = $_.FullName.Substring($base.Length).TrimStart([char[]]'\/').Replace('\','/')
    $manifest[$relative] = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
}
[IO.File]::WriteAllText((Join-Path $Output 'sha256.json'), ($manifest | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
Write-Host 'PAQUETE PREPARADO:' $Output
Write-Host 'Todavia no se ha modificado el servicio instalado.'
