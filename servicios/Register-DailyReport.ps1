[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$RepositoryPath,
    [Parameter(Mandatory)][string]$PythonPath,
    [string]$BaseDir = 'C:\Rentabilidad',
    [string]$SqlConfigPath = 'C:\Rentabilidad\sql_config.json',
    [string]$TaskName = 'Rentabilidad-Diaria-SQL'
)

$ErrorActionPreference = 'Stop'
if ((Get-TimeZone).Id -ne 'SA Pacific Standard Time') {
    throw 'La tarea debe registrarse con el servidor en la zona horaria de Colombia (SA Pacific Standard Time).'
}
foreach ($required in @($RepositoryPath, $PythonPath, $SqlConfigPath, "$BaseDir\PLANTILLA.xlsx")) {
    if (-not (Test-Path -LiteralPath $required)) { throw "No existe: $required" }
}
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    throw "Ya existe la tarea $TaskName. Revísala antes de sustituirla."
}
$arguments = '-m rentabilidad.daily_report --base-dir "{0}" --sql-config "{1}"' -f $BaseDir, $SqlConfigPath
$action = New-ScheduledTaskAction -Execute $PythonPath -Argument $arguments -WorkingDirectory $RepositoryPath
$trigger = New-ScheduledTaskTrigger -Daily -At '05:00'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 15) `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 45)
if ($PSCmdlet.ShouldProcess($TaskName, 'Registrar informe diario SQL a las 05:00')) {
    # Las credenciales se introducen en Windows, nunca como argumentos ni en Git.
    $credential = Get-Credential -Message 'Cuenta Windows para ejecutar el informe sin una sesión abierta'
    if ($null -eq $credential) { throw 'No se proporcionó una cuenta de ejecución.' }
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
        -Settings $settings -User $credential.UserName `
        -Password $credential.GetNetworkCredential().Password `
        -Description 'Informe del día anterior desde SQL; 05:00 hora de Colombia.' | Out-Null
    Write-Output "Registrada: $TaskName. Verifique el primer resultado y el log en $BaseDir\Logs."
}
