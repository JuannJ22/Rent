[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][datetime]$Desde,
    [Parameter(Mandatory=$true)][datetime]$Hasta,
    [switch]$Replace
)
$ErrorActionPreference = 'Stop'
if ($Hasta.Date -lt $Desde.Date) { throw 'Hasta no puede ser anterior a Desde.' }
$lastAllowed = (Get-Date).ToUniversalTime().AddHours(-5).Date.AddDays(-1)
if ($Hasta.Date -gt $lastAllowed) { throw 'El lote historico llega como maximo hasta ayer en Colombia.' }
$results = @()
for ($day = $Desde.Date; $day -le $Hasta.Date; $day = $day.AddDays(1)) {
    $date = $day.ToString('yyyy-MM-dd')
    $files = @(Get-ChildItem 'C:\Rentabilidad\Productos' -File -Filter '*.xlsx' | ForEach-Object {
        $label = $null
        if ($_.BaseName -match '^productos-(\d{4}-\d{2}-\d{2})$') { $label = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', $null) }
        elseif ($_.BaseName -match '^productos(\d{4})$') { $label = [datetime]::ParseExact(($Desde.Year.ToString() + $Matches[1]), 'yyyyMMdd', $null) }
        if ($label) { [pscustomobject]@{ Fecha=$label; Ruta=$_.FullName } }
    } | Sort-Object Fecha)
    if ($files.Count -eq 0) { throw 'No se encontraron listados de productos.' }
    $selected = @($files | Where-Object Fecha -le $day | Select-Object -Last 1)
    if ($selected.Count -eq 0) { $selected = @($files | Select-Object -First 1) }
    $previousProducts = $env:SQL_LEGACY_PRODUCT_FILE
    $env:SQL_LEGACY_PRODUCT_FILE = $selected[0].Ruta
    try {
        $arguments = @{ Fecha=$date }
        if ($Replace) { $arguments.Replace = $true }
        & 'C:\Rentabilidad\ActualizarETL\Test-Rentabilidad.ps1' @arguments
        $results += [pscustomobject]@{ Fecha=$date; Estado='CORRECTO'; Productos=$selected[0].Ruta; Error='' }
    }
    catch {
        $results += [pscustomobject]@{ Fecha=$date; Estado='ERROR'; Productos=$selected[0].Ruta; Error=$_.Exception.Message }
    }
    finally { $env:SQL_LEGACY_PRODUCT_FILE = $previousProducts }
}
$summary = 'C:\Rentabilidad\lote-' + $Desde.ToString('yyyyMMdd') + '-' + $Hasta.ToString('yyyyMMdd') + '-' + (Get-Date -Format 'yyyyMMdd-HHmmss-ffff') + '.csv'
$results | Export-Csv $summary -NoTypeInformation -Encoding UTF8
$results | Format-Table Fecha, Estado, Error -Wrap
Write-Host 'Resumen:' $summary
if (@($results | Where-Object Estado -ne 'CORRECTO').Count -gt 0) {
    throw 'El lote tiene fechas pendientes. Los informes correctos se conservaron; consulta el resumen.'
}
