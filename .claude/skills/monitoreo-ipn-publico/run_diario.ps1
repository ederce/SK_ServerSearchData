<#
Wrapper para Task Scheduler: corre monitorear_ipn_publico.py con la fecha
de AYER (calculada al momento de ejecutarse), sin intervencion manual.
Requiere que sftp_config.json (crypto/ de la skill hermana
busqueda-transacciones-local-sftp) ya tenga host/usuario/password
guardados -- si faltan, el script pediria input() y la tarea programada
se quedaria colgada esperando una consola que no existe.
#>

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = "C:\Users\MC2343\AppData\Local\Programs\Python\Python312\python.exe"
$fecha = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")

$logDir = Join-Path $scriptDir "logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}
$logFile = Join-Path $logDir "run_$fecha.log"

Set-Location $scriptDir
& $python "monitorear_ipn_publico.py" --fecha $fecha *>> $logFile
$exitCode = $LASTEXITCODE

"Salida: $exitCode" | Add-Content -Path $logFile
exit $exitCode
